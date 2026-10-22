using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace PDMReReferenceAddin
{
    [ComVisible(true)]
    [Guid("A6A1A0D4-76E2-4D31-9C0D-3A2F1B8C7E91")]
    public class ReReferenceAddin : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo info, IEdmVault5 vault, IEdmCmdMgr5 cmdMgr)
        {
            info.mbsAddInName = "Re-Reference Automation";
            info.mbsCompany = "Internal Tools";
            info.mbsDescription = "Single-file test and Excel batch re-reference automation";
            info.mlAddInVersion = 3;
            info.mlRequiredVersionMajor = 5;
            info.mlRequiredVersionMinor = 2;

            cmdMgr.AddCmd(1000, "Batch Re-Reference...", (int)EdmMenuFlags.EdmMenu_Nothing,
                "Open the single-file test or Excel batch form", "Batch Re-Reference");
        }

        public void OnCmd(ref EdmCmd cmd, ref EdmCmdData[] data)
        {
            if (cmd.meCmdType != EdmCmdType.EdmCmd_Menu || cmd.mlCmdID != 1000) return;
            try
            {
                using (ReReferenceForm form = new ReReferenceForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                        RunDryRun(form, cmd.mpoVault as IEdmVault5);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Batch Re-Reference", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RunDryRun(ReReferenceForm form, IEdmVault5 vault)
        {
            if (form.Mode == ReReferenceMode.ExcelBatch)
            {
                List<BatchPlanRow> rows = BatchPlanReader.ReadWorkbook(form.WorkbookPath);
                string logPath = Path.Combine(Path.GetDirectoryName(form.WorkbookPath),
                    "ReReference_RunLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
                BatchRunCoordinator coordinator = new BatchRunCoordinator();
                coordinator.ProgressChanged += delegate(object sender, BatchProgressEventArgs e)
                {
                    // This initial version is intentionally synchronous and dry-run only.
                };
                IList<BatchRowResult> results;
                if (vault == null)
                {
                    MessageBox.Show("The invoking PDM vault session was not available. Run this command from inside the pdmdev vault.",
                        "PDM session unavailable", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                results = coordinator.Run(rows, new PdmDryRunRowProcessor(vault), logPath, CancellationToken.None);
                int inspected = 0, review = 0, validatedOnly = 0;
                foreach (BatchRowResult result in results)
                {
                    if (result.Status == "DRY_RUN_INSPECTED") inspected++;
                    else if (result.Status == "DRY_RUN_ONLY") validatedOnly++;
                    else if (result.Status == "REVIEW") review++;
                }
                MessageBox.Show("Excel PDM dry run complete.\r\n\r\nRows inspected: " + inspected +
                    "\r\nRows needing review: " + review +
                    "\r\nRows validated without PDM session: " + validatedOnly + "\r\n\r\nLog:\r\n" + logPath +
                    "\r\n\r\nRead-only inspection only. No PDM files were changed.", "Batch Re-Reference",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string logDirectory = Path.GetDirectoryName(form.SinglePartPath);
                string logPath = Path.Combine(logDirectory,
                    "ReReference_SingleFile_DryRun_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv");
                string sourceId = "";
                string sourceVaultPath = "";
                string replacementId = "";
                string replacementVaultPath = "";
                string sourceStatus = "NOT_CHECKED";
                string replacementStatus = "NOT_CHECKED";
                string message = "";

                // Read-only validation against the vault session that invoked this add-in.
                // No checkout, reference edit, state transition, check-in, or rename is performed.
                try
                {
                    if (vault == null)
                    {
                        message = "PDM vault session was not available to the add-in command.";
                    }
                    else
                    {
                        IEdmFolder5 sourceFolder;
                        IEdmFile5 sourceFile = vault.GetFileFromPath(form.SinglePartPath, out sourceFolder);
                        if (sourceFile != null)
                        {
                            sourceId = sourceFile.ID.ToString();
                            sourceVaultPath = form.SinglePartPath;
                            sourceStatus = "RESOLVED";
                        }
                        else
                        {
                            sourceStatus = "NOT_RESOLVED";
                            message = "Source path did not resolve to a PDM file object.";
                        }

                        string replacementPath = Path.Combine(Path.GetDirectoryName(form.SinglePartPath), form.ReplacementPartName);
                        IEdmFolder5 replacementFolder;
                        IEdmFile5 replacementFile = vault.GetFileFromPath(replacementPath, out replacementFolder);
                        if (replacementFile != null)
                        {
                            replacementId = replacementFile.ID.ToString();
                            replacementVaultPath = replacementPath;
                            replacementStatus = "RESOLVED";
                        }
                        else
                        {
                            replacementStatus = "NOT_RESOLVED";
                            if (!String.IsNullOrEmpty(message)) message += " ";
                            message += "Replacement filename did not resolve in the source file's folder.";
                        }
                    }
                }
                catch (Exception validationException)
                {
                    if (!String.IsNullOrEmpty(message)) message += " ";
                    message += "PDM validation error: " + validationException.Message;
                }

                using (StreamWriter writer = new StreamWriter(logPath, false, new UTF8Encoding(true)))
                {
                    writer.WriteLine("Timestamp,Mode,SourceLocalPath,SourcePdmFileId,SourceVaultPath,SourceStatus,ReplacementFilename,ReplacementPdmFileId,ReplacementVaultPath,ReplacementStatus,ParentAssemblies,WorkflowStates,ExecutionMode,Message");
                    writer.WriteLine(String.Join(",", Csv(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                        Csv("Single File PDM Validation"), Csv(form.SinglePartPath), Csv(sourceId), Csv(sourceVaultPath), Csv(sourceStatus),
                        Csv(form.ReplacementPartName), Csv(replacementId), Csv(replacementVaultPath), Csv(replacementStatus),
                        Csv("NOT_CHECKED"), Csv("NOT_CHECKED"), Csv("DRY_RUN_ONLY"),
                        Csv(String.IsNullOrEmpty(message) ? "Read-only object resolution only. No PDM data changed." : message + " No PDM data changed.")));
                }

                MessageBox.Show("Single-file PDM validation complete.\r\n\r\nSource: " + sourceStatus +
                    (String.IsNullOrEmpty(sourceId) ? "" : " (ID " + sourceId + ")") +
                    "\r\nReplacement: " + replacementStatus +
                    (String.IsNullOrEmpty(replacementId) ? "" : " (ID " + replacementId + ")") +
                    "\r\nParent assemblies: NOT_CHECKED\r\nWorkflow states: NOT_CHECKED\r\n\r\nLog:\r\n" + logPath +
                    "\r\n\r\nDRY RUN ONLY — no PDM data was changed.\r\n" + message, "Batch Re-Reference",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string Csv(string value)
        {
            value = value ?? "";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }

    public enum ReReferenceMode { SingleFile, ExcelBatch }

    public sealed class ReReferenceForm : Form
    {
        private readonly RadioButton singleMode;
        private readonly RadioButton excelMode;
        private readonly TextBox singlePartText;
        private readonly TextBox replacementText;
        private readonly TextBox workbookText;
        private readonly Button singleBrowseButton;
        private readonly Button excelBrowseButton;
        private readonly Button runButton;
        private readonly Button cancelButton;
        private readonly Label singleHint;
        private readonly Label excelHint;
        private readonly Label singleLabel;
        private readonly Label replacementLabel;
        private readonly Label excelLabel;

        public ReReferenceMode Mode { get { return excelMode.Checked ? ReReferenceMode.ExcelBatch : ReReferenceMode.SingleFile; } }
        public string SinglePartPath { get { return singlePartText.Text.Trim(); } }
        public string ReplacementPartName { get { return replacementText.Text.Trim(); } }
        public string WorkbookPath { get { return workbookText.Text.Trim(); } }

        public ReReferenceForm()
        {
            Text = "Batch Re-Reference";
            Width = 690;
            Height = 440;
            MinimumSize = new System.Drawing.Size(650, 400);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            Label heading = new Label { Text = "PDM Re-Reference Tool", Left = 20, Top = 16, Width = 620, Height = 28, Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold) };
            Label intro = new Label { Text = "Integrated read-only PDM inspection. No checkout, workflow changes, reference edits, check-in, or rename.", Left = 22, Top = 50, Width = 630, Height = 35 };
            singleMode = new RadioButton { Text = "Single File Test", Left = 24, Top = 96, Width = 180, Checked = true, AutoSize = true };
            excelMode = new RadioButton { Text = "Excel Batch Run", Left = 230, Top = 96, Width = 180, AutoSize = true };

            singleLabel = new Label { Text = "Part file (.sldprt or .sldasm)", Left = 24, Top = 135, Width = 300 };
            singlePartText = new TextBox { Left = 24, Top = 158, Width = 490 };
            singleBrowseButton = new Button { Text = "Browse...", Left = 526, Top = 156, Width = 110, Height = 27 };
            replacementLabel = new Label { Text = "Replacement part filename (for recording the test intent)", Left = 24, Top = 194, Width = 500 };
            replacementText = new TextBox { Left = 24, Top = 217, Width = 612 };
            singleHint = new Label { Text = "Note: file browsing here selects a local file; it does not yet resolve the PDM vault object or change references.", Left = 24, Top = 249, Width = 620, Height = 34, ForeColor = System.Drawing.Color.DarkSlateGray };

            excelLabel = new Label { Text = "Excel workbook (.xlsx)", Left = 24, Top = 135, Width = 300 };
            workbookText = new TextBox { Left = 24, Top = 158, Width = 490 };
            excelBrowseButton = new Button { Text = "Browse...", Left = 526, Top = 156, Width = 110, Height = 27 };
            excelHint = new Label { Text = "The workbook is validated row by row and a CSV log is created. No PDM files are changed.", Left = 24, Top = 194, Width = 620, Height = 34, ForeColor = System.Drawing.Color.DarkSlateGray };

            runButton = new Button { Text = "Run Dry Run", Left = 394, Top = 330, Width = 120, Height = 34, DialogResult = DialogResult.None };
            cancelButton = new Button { Text = "Cancel", Left = 526, Top = 330, Width = 110, Height = 34, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { heading, intro, singleMode, excelMode, singleLabel, singlePartText, singleBrowseButton, replacementLabel, replacementText, singleHint, excelLabel, workbookText, excelBrowseButton, excelHint, runButton, cancelButton });
            AcceptButton = runButton;
            CancelButton = cancelButton;

            singleBrowseButton.Click += delegate
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Title = "Select a part or assembly file";
                    dialog.Filter = "SOLIDWORKS files (*.sldprt;*.sldasm)|*.sldprt;*.sldasm|All files (*.*)|*.*";
                    dialog.CheckFileExists = true;
                    if (dialog.ShowDialog(this) == DialogResult.OK) singlePartText.Text = dialog.FileName;
                }
            };
            excelBrowseButton.Click += delegate
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Title = "Select the re-reference Excel workbook";
                    dialog.Filter = "Excel workbook (*.xlsx)|*.xlsx";
                    dialog.CheckFileExists = true;
                    if (dialog.ShowDialog(this) == DialogResult.OK) workbookText.Text = dialog.FileName;
                }
            };
            singleMode.CheckedChanged += delegate { UpdateMode(); };
            excelMode.CheckedChanged += delegate { UpdateMode(); };
            singleMode.Click += delegate { UpdateMode(); };
            excelMode.Click += delegate { UpdateMode(); };
            runButton.Click += delegate { ValidateAndAccept(); };
            UpdateMode();
        }

        private void UpdateMode()
        {
            bool single = singleMode.Checked;

            singleLabel.Visible = single;
            singleLabel.Enabled = single;
            singlePartText.Visible = single;
            singlePartText.Enabled = single;
            singleBrowseButton.Visible = single;
            singleBrowseButton.Enabled = single;
            replacementLabel.Visible = single;
            replacementLabel.Enabled = single;
            replacementText.Visible = single;
            replacementText.Enabled = single;
            singleHint.Visible = single;
            singleHint.Enabled = single;

            excelLabel.Visible = !single;
            excelLabel.Enabled = !single;
            workbookText.Visible = !single;
            workbookText.Enabled = !single;
            excelBrowseButton.Visible = !single;
            excelBrowseButton.Enabled = !single;
            excelHint.Visible = !single;
            excelHint.Enabled = !single;

            runButton.Text = single ? "Run Single Test" : "Run Excel Batch";
        }



        private void ValidateAndAccept()
        {
            if (Mode == ReReferenceMode.SingleFile)
            {
                if (String.IsNullOrWhiteSpace(SinglePartPath) || !File.Exists(SinglePartPath))
                {
                    MessageBox.Show(this, "Select an existing local part or assembly file.", "Part required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (String.IsNullOrWhiteSpace(ReplacementPartName))
                {
                    MessageBox.Show(this, "Enter the replacement part filename so the test intent is recorded.", "Replacement required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else if (String.IsNullOrWhiteSpace(WorkbookPath) || !File.Exists(WorkbookPath))
            {
                MessageBox.Show(this, "Select an existing Excel workbook (.xlsx).", "Workbook required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
