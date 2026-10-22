using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using EPDM.Interop.epdm;

namespace PDMReReferenceAddin
{
    /// <summary>
    /// Read-only PDM inspection for the integrated dry run.
    /// Deliberately contains no checkout, transition, rename, check-in, or reference-edit calls.
    /// Uses late binding for reference-tree/search members whose interop signatures differ by PDM version.
    /// </summary>
    public sealed class PdmDryRunRowProcessor : IReReferenceRowProcessor
    {
        private readonly IEdmVault5 vault;

        public PdmDryRunRowProcessor(IEdmVault5 vaultSession)
        {
            vault = vaultSession ?? throw new ArgumentNullException("vaultSession");
        }

        public BatchRowResult Process(BatchPlanRow row, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            BatchRowResult result = new BatchRowResult
            {
                ExcelRowNumber = row.ExcelRowNumber,
                SourcePart = row.SourcePart,
                ReplacementPart = row.ReplacementPart,
                StartedAt = DateTime.Now
            };

            if (!String.Equals(row.ValidationStatus, "VALID", StringComparison.OrdinalIgnoreCase))
            {
                result.Status = "REVIEW";
                result.Message = String.IsNullOrWhiteSpace(row.ValidationMessage)
                    ? "Workbook row did not pass validation."
                    : row.ValidationMessage;
                result.FinishedAt = DateTime.Now;
                return result;
            }

            try
            {
                FileMatch source = FindUniqueFile(row.SourcePart);
                FileMatch replacement = FindUniqueFile(row.ReplacementPart);
                List<string> notes = new List<string>();
                notes.Add("Source=" + source.Status + (source.Id > 0 ? " ID=" + source.Id : ""));
                notes.Add("SourcePath=" + (source.Path ?? ""));
                notes.Add("Replacement=" + replacement.Status + (replacement.Id > 0 ? " ID=" + replacement.Id : ""));
                notes.Add("ReplacementPath=" + (replacement.Path ?? ""));
                notes.Add("SolidWorksPreflight=" + InspectSolidWorksAvailability());
                notes.Add("SourceLocalFile=" + InspectLocalFile(source.Path));
                notes.Add("ReplacementLocalFile=" + InspectLocalFile(replacement.Path));

                if (source.Status != "RESOLVED_UNIQUE" || replacement.Status != "RESOLVED_UNIQUE")
                {
                    result.Status = "REVIEW";
                    notes.Add("Resolve names to exactly one PDM file before any future execution.");
                }
                else if (source.Id == replacement.Id)
                {
                    result.Status = "REVIEW";
                    notes.Add("Source and replacement resolve to the same PDM file ID.");
                }
                else
                {
                    ParentInspection parents = InspectParents(source);
                    notes.Add("ParentAssemblies=" + parents.AssemblyCount);
                    notes.Add("DrawingsIgnored=" + parents.DrawingCount);
                    notes.Add("ParentDetails=" + (parents.Details.Count == 0 ? "(none found)" : String.Join(" || ", parents.Details.ToArray())));
                    notes.Add("WorkflowPlanSummary=" + parents.WorkflowPlanSummary);
                    notes.Add("WorkflowPlanBasis=VBA state-name rules only; actual transition availability and editability are NOT yet verified. CheckoutPreflight is read-only metadata only.");
                    notes.Add("PreflightScope=read-only; local file existence and running SOLIDWORKS availability are observations only.");
                    notes.Add("ExpectedAssemblyCount=" + row.ExpectedAssemblyCount);
                    if (!String.IsNullOrEmpty(parents.Warning)) notes.Add("ReferenceTreeWarning=" + parents.Warning);
                    if (parents.UnsupportedStateCount > 0)
                        notes.Add("UnsupportedWorkflowStates=" + parents.UnsupportedStateCount + " (manual review required before any write operation)");
                    if (parents.Warning.Length > 0 || parents.UnsupportedStateCount > 0)
                        result.Status = "REVIEW";
                    else
                        result.Status = "DRY_RUN_INSPECTED";
                }
                result.Message = String.Join("; ", notes.ToArray()) + "; READ_ONLY: no PDM changes made.";
            }
            catch (Exception ex)
            {
                result.Status = "REVIEW";
                result.Message = "Read-only PDM inspection failed: " + ex.GetType().Name + ": " + ex.Message;
            }

            result.FinishedAt = DateTime.Now;
            return result;
        }

        private FileMatch FindUniqueFile(string filename)
        {
            string name = Path.GetFileName(filename ?? "").Trim();
            if (name.Length == 0) return new FileMatch { Status = "NOT_FOUND" };

            List<FileMatch> matches = new List<FileMatch>();
            if (Path.HasExtension(name))
            {
                SearchExactName(name, matches);
            }
            else
            {
                // The workbook contains filename stems without extensions.
                // Search the two SOLIDWORKS document types expected by this macro.
                SearchExactName(name + ".SLDPRT", matches);
                SearchExactName(name + ".SLDASM", matches);
            }

            // De-duplicate identical results if the PDM search API returns the same object twice.
            Dictionary<int, FileMatch> unique = new Dictionary<int, FileMatch>();
            foreach (FileMatch item in matches)
                if (item.Id > 0 && !unique.ContainsKey(item.Id)) unique.Add(item.Id, item);

            if (unique.Count == 0) return new FileMatch { Status = "NOT_FOUND" };
            if (unique.Count > 1) return new FileMatch { Status = "AMBIGUOUS_" + unique.Count };
            foreach (FileMatch item in unique.Values) return item;
            return new FileMatch { Status = "NOT_FOUND" };
        }

        private void SearchExactName(string name, List<FileMatch> matches)
        {
            dynamic search = ((dynamic)vault).CreateSearch();
            search.FileName = name;
            search.FindFiles = true;
            search.FindFolders = false;
            search.Recursive = true;

            dynamic result = search.GetFirstResult();
            int guard = 0;
            while (result != null && guard++ < 1000)
            {
                matches.Add(new FileMatch
                {
                    Id = Convert.ToInt32(result.ID),
                    Path = Convert.ToString(result.Path),
                    Status = "RESOLVED_UNIQUE"
                });
                // Some PDM COM search implementations signal end-of-results by throwing
                // a COMException rather than returning null. Treat that as normal search
                // termination; do not let an exhausted search cursor fail the row.
                try
                {
                    result = search.GetNextResult();
                }
                catch (COMException)
                {
                    break;
                }
            }
            if (guard >= 1000)
                throw new InvalidOperationException("PDM search safety limit reached for " + name + ".");
        }

        private ParentInspection InspectParents(FileMatch source)
        {
            ParentInspection inspection = new ParentInspection();
            IEdmFolder5 sourceFolder;
            IEdmFile5 sourceFile = vault.GetFileFromPath(source.Path, out sourceFolder);
            if (sourceFile == null || sourceFolder == null)
            {
                inspection.Warning = "Resolved source path could not be reopened as a PDM file/folder.";
                return inspection;
            }

            // Use the installed PDM interop interfaces directly instead of late-bound COM calls.
            // GetReferenceTree's second parameter (version number) is optional and defaults to 0.
            IEdmReference5 referenceTree = sourceFile.GetReferenceTree(sourceFolder.ID, 0);
            if (referenceTree == null)
            {
                inspection.Warning = "GetReferenceTree returned no reference tree.";
                return inspection;
            }

            IEdmReference7 referenceTree7 = referenceTree as IEdmReference7;
            IEdmPos5 position;
            try
            {
                position = referenceTree7 != null
                    ? referenceTree7.GetFirstParentPosition2(0, false, 15)
                    : referenceTree.GetFirstParentPosition(0, false);
            }
            catch (COMException ex)
            {
                inspection.Warning = "Could not initialize parent reference cursor: " + ex.Message;
                return inspection;
            }

            HashSet<int> seen = new HashSet<int>();
            int guard = 0;
            while (position != null && guard++ < 10000)
            {
                if (position.IsNull) break;

                IEdmReference5 parentRef;
                try
                {
                    // GetNextParent accepts the current IEdmPos5 as an input parameter;
                    // the COM reference-tree object advances its internal enumeration state.
                    parentRef = referenceTree.GetNextParent(position);
                }
                catch (COMException ex)
                {
                    inspection.Warning = "Parent reference enumeration failed before completion: " + ex.Message;
                    break;
                }
                if (parentRef == null) break;

                IEdmFile5 parentFile;
                try { parentFile = parentRef.File; }
                catch (COMException ex)
                {
                    inspection.Warning = "Could not read a parent file from the reference tree: " + ex.Message;
                    break;
                }
                if (parentFile == null) continue;

                int parentId = Convert.ToInt32(parentFile.ID);
                if (parentId == source.Id || !seen.Add(parentId)) continue;
                string parentName = Convert.ToString(parentFile.Name) ?? "";
                string extension = Path.GetExtension(parentName);
                if (String.Equals(extension, ".slddrw", StringComparison.OrdinalIgnoreCase))
                {
                    inspection.DrawingCount++;
                    continue;
                }
                if (!String.Equals(extension, ".sldasm", StringComparison.OrdinalIgnoreCase)) continue;

                inspection.AssemblyCount++;
                string state = "(state unavailable)";
                try
                {
                    IEdmState5 currentState = parentFile.CurrentState;
                    if (currentState != null) state = Convert.ToString(currentState.Name) ?? state;
                }
                catch { }
                string parentPath = GetLocalPath(parentFile);
                string localFilePreflight = InspectLocalFile(parentPath);
                string checkoutPreflight = InspectCheckoutReadOnly(parentFile);
                string workflowPlan = GetWorkflowPlan(state, out bool supportedState);
                string transitionInfo = PdmTransitionInspector.Inspect(parentFile);
                if (!supportedState) inspection.UnsupportedStateCount++;
                inspection.WorkflowPlanCounts[workflowPlan] = inspection.WorkflowPlanCounts.ContainsKey(workflowPlan)
                    ? inspection.WorkflowPlanCounts[workflowPlan] + 1 : 1;
                inspection.Details.Add(parentName + " [ID=" + parentId + "; State=" + state + "; LocalFilePreflight=" + localFilePreflight + "; CheckoutPreflight=" + checkoutPreflight + "; PlannedWorkflow=" + workflowPlan + "; Path=" + parentPath + "; " + transitionInfo + "]");
            }
            if (guard >= 10000 && String.IsNullOrEmpty(inspection.Warning)) inspection.Warning = "Parent enumeration safety limit reached.";
            inspection.WorkflowPlanSummary = inspection.WorkflowPlanCounts.Count == 0
                ? "(no parent assembly workflow plans)"
                : String.Join(", ", BuildWorkflowSummary(inspection.WorkflowPlanCounts));
            return inspection;
        }



        // Read-only probe: asks the Windows COM Running Object Table for a running SOLIDWORKS instance.
        private static string InspectSolidWorksAvailability()
        {
            object sw = null;
            try
            {
                Type t = Type.GetTypeFromProgID("SldWorks.Application");
                if (t == null) return "PROGID_NOT_REGISTERED";
                try { sw = Marshal.GetActiveObject("SldWorks.Application"); }
                catch (COMException) { return "NOT_RUNNING_OR_NOT_REGISTERED_IN_ROT"; }
                if (sw == null) return "NOT_RUNNING_OR_NOT_REGISTERED_IN_ROT";
                try { return "RUNNING;Revision=" + Convert.ToString(sw.GetType().InvokeMember("RevisionNumber", BindingFlags.GetProperty, null, sw, null)); }
                catch { return "RUNNING;REVISION_UNAVAILABLE"; }
            }
            catch (Exception ex) { return "PROBE_ERROR=" + ex.GetType().Name + ":" + ex.Message; }
            finally { if (sw != null && Marshal.IsComObject(sw)) Marshal.ReleaseComObject(sw); }
        }

        private static string InspectLocalFile(string path)
        {
            if (String.IsNullOrWhiteSpace(path) || path.StartsWith("(", StringComparison.Ordinal)) return "LOCAL_PATH_UNAVAILABLE";
            try { return File.Exists(path) ? "EXISTS" : "MISSING_OR_NOT_CACHED"; }
            catch (Exception ex) { return "CHECK_ERROR=" + ex.GetType().Name + ":" + ex.Message; }
        }

        // Read-only checkout metadata probe. It never unlocks or checks out a file.
        private static string InspectCheckoutReadOnly(IEdmFile5 file)
        {
            try
            {
                dynamic d = file;
                bool isLocked = false;
                try { isLocked = Convert.ToBoolean(d.IsLocked); }
                catch
                {
                    try { isLocked = Convert.ToInt32(d.LockedByUserID) > 0; }
                    catch { return "LOCK_STATUS_UNAVAILABLE"; }
                }

                if (!isLocked) return "NOT_CHECKED_OUT";
                int lockedBy = 0;
                try { lockedBy = Convert.ToInt32(d.LockedByUserID); }
                catch { return "CHECKED_OUT_OWNER_UNKNOWN"; }

                string ownerName = "";
                try
                {
                    dynamic userMgr = ((dynamic)file).Vault.CreateUtility(3);
                    dynamic user = userMgr.GetUser(lockedBy);
                    if (user != null) ownerName = Convert.ToString(user.Name) ?? "";
                }
                catch { }

                return "CHECKED_OUT_BY_ID=" + lockedBy + (ownerName.Length > 0 ? ";OWNER=" + ownerName : ";OWNER_NAME_UNAVAILABLE");
            }
            catch (Exception ex)
            {
                return "PREFLIGHT_ERROR=" + ex.GetType().Name + ":" + ex.Message;
            }
        }

        // Mirrors the state-name decision branches in BringToWip/ReturnToReleased in the VBA.
        // This is a planning label only; it does not check transition availability or change state.
        private static string GetWorkflowPlan(string stateName, out bool supported)
        {
            string state = (stateName ?? "").Trim().ToLowerInvariant();
            bool isWip = state == "wip" || (state.Contains("work") && state.Contains("progress"));
            bool isReleasedManufacturing = state.Contains("released") && state.Contains("manufactur");
            bool isUnderReview = state.Contains("under") && state.Contains("review");
            bool isApproved = state.Contains("approved") && !isUnderReview;

            if (isWip)
            {
                supported = true;
                return "WIP_NO_STATE_TRANSITION_PLANNED";
            }
            if (isReleasedManufacturing)
            {
                supported = true;
                return "RFM_TO_REVIEW_TO_APPROVED_TO_WIP_THEN_RETURN_TO_RELEASED";
            }
            if (isUnderReview)
            {
                supported = true;
                return "REVIEW_TO_APPROVED_TO_WIP";
            }
            if (isApproved)
            {
                supported = true;
                return "APPROVED_TO_WIP_IF_OUTGOING_TRANSITION_EXISTS";
            }

            supported = false;
            return "UNSUPPORTED_STATE_MANUAL_REVIEW";
        }

        private static List<string> BuildWorkflowSummary(Dictionary<string, int> counts)
        {
            List<string> parts = new List<string>();
            foreach (KeyValuePair<string, int> item in counts)
                parts.Add(item.Key + "=" + item.Value);
            parts.Sort(StringComparer.Ordinal);
            return parts;
        }

        private static string GetLocalPath(dynamic file)
        {
            try
            {
                dynamic folderPosition = file.GetFirstFolderPosition();
                int guard = 0;
                while (folderPosition != null && guard++ < 1000)
                {
                    bool isNull = false;
                    try { isNull = Convert.ToBoolean(folderPosition.IsNull); } catch { }
                    if (isNull) break;
                    dynamic folder = file.GetNextFolder(folderPosition);
                    if (folder == null) break;
                    string path = Convert.ToString(file.GetLocalPath(Convert.ToInt32(folder.ID)));
                    if (!String.IsNullOrWhiteSpace(path)) return path;
                }
            }
            catch { }
            return "(local path unavailable)";
        }

        private sealed class FileMatch
        {
            public string Status;
            public int Id;
            public string Path;
        }

        private sealed class ParentInspection
        {
            public int AssemblyCount;
            public int DrawingCount;
            public int UnsupportedStateCount;
            public string Warning = "";
            public string WorkflowPlanSummary = "";
            public readonly Dictionary<string, int> WorkflowPlanCounts = new Dictionary<string, int>();
            public readonly List<string> Details = new List<string>();
        }
    }
}
