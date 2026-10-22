using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace PDMReReferenceAddin
{
    // Reads the first worksheet of an .xlsx workbook without requiring
    // Excel or a third-party NuGet package. Designed for .NET Framework 4.7.2.
    public sealed class BatchPlanRow
    {
        public int ExcelRowNumber { get; set; }
        public string PartName { get; set; }
        public string CopyVersionName { get; set; }
        public int ExpectedAssemblyCount { get; set; }
        public string PartToReference { get; set; }

        // The macro's operation is: replace SourcePart with ReplacementPart
        // in parent assemblies. The workbook's last column identifies the
        // desired replacement; the other member of the pair is the source.
        public string SourcePart { get; set; }
        public string ReplacementPart { get; set; }
        public string ValidationStatus { get; set; }
        public string ValidationMessage { get; set; }
    }

    public static class BatchPlanReader
    {
        private static readonly XNamespace MainNs =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace RelNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        public static List<BatchPlanRow> ReadWorkbook(string path)
        {
            if (String.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Select an Excel workbook.", "path");
            if (!File.Exists(path))
                throw new FileNotFoundException("Workbook was not found.", path);
            if (!String.Equals(Path.GetExtension(path), ".xlsx",
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "Please select an .xlsx workbook. Legacy .xls is not supported.");

            using (FileStream fileStream = File.OpenRead(path))
            using (ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Read, false))
            {
                XDocument workbook = LoadXml(archive, "xl/workbook.xml");
                XDocument workbookRels = LoadXml(archive, "xl/_rels/workbook.xml.rels");
                string firstSheetRelId = (string)workbook
                    .Descendants(MainNs + "sheet").FirstOrDefault()?
                    .Attribute(RelNs + "id");

                if (String.IsNullOrWhiteSpace(firstSheetRelId))
                    throw new InvalidDataException("The workbook has no worksheet.");

                XElement relation = workbookRels.Root.Elements(PackageRelNs + "Relationship")
                    .FirstOrDefault(x => (string)x.Attribute("Id") == firstSheetRelId);

                if (relation == null)
                    throw new InvalidDataException(
                        "Could not resolve the first worksheet in the workbook.");

                string target = (string)relation.Attribute("Target");
                string sheetPath = NormalizeWorksheetPath(target);
                XDocument sheet = LoadXml(archive, sheetPath);
                List<string> sharedStrings = ReadSharedStrings(archive);
                List<List<string>> grid = ReadGrid(sheet, sharedStrings);

                if (grid.Count == 0)
                    throw new InvalidDataException("The first worksheet is empty.");

                List<string> headers = grid[0].Select(x => (x ?? "").Trim()).ToList();
                int partNameCol = FindHeader(headers, "Name of the Part");
                int copyNameCol = FindHeader(headers, "Copy Version Name");
                int assemblyCountCol = FindHeader(headers, "No. of assemblies");
                int targetCol = FindHeader(headers, "Part which we need to refer to");

                if (partNameCol < 0 || copyNameCol < 0 ||
                    assemblyCountCol < 0 || targetCol < 0)
                {
                    throw new InvalidDataException(
                        "Required headers were not found in the first worksheet. " +
                        "Expected: Name of the Part, Copy Version Name, " +
                        "No. of assemblies, Part which we need to refer to.");
                }

                List<BatchPlanRow> rows = new List<BatchPlanRow>();
                for (int i = 1; i < grid.Count; i++)
                {
                    List<string> cells = grid[i];
                    string partName = Cell(cells, partNameCol).Trim();
                    string copyName = Cell(cells, copyNameCol).Trim();
                    string countText = Cell(cells, assemblyCountCol).Trim();
                    string targetName = Cell(cells, targetCol).Trim();

                    // Ignore blank rows and summary/footer rows. The supplied
                    // workbook has a total row with only the assembly count (65).
                    // It is not a part mapping and must not be sent for processing.
                    if (partName.Length == 0 && copyName.Length == 0 &&
                        targetName.Length == 0)
                        continue;

                    BatchPlanRow row = new BatchPlanRow
                    {
                        ExcelRowNumber = i + 1,
                        PartName = partName,
                        CopyVersionName = copyName,
                        PartToReference = targetName,
                        ValidationStatus = "VALID",
                        ValidationMessage = ""
                    };

                    int expected;
                    if (!Int32.TryParse(countText, out expected) || expected < 0)
                    {
                        row.ValidationStatus = "REVIEW";
                        row.ValidationMessage = "Assembly count is blank or not a non-negative integer.";
                    }
                    else
                    {
                        row.ExpectedAssemblyCount = expected;
                    }

                    if (partName.Length == 0 || copyName.Length == 0 || targetName.Length == 0)
                    {
                        row.ValidationStatus = "REVIEW";
                        row.ValidationMessage = Append(row.ValidationMessage,
                            "One or more required part-name cells are blank.");
                    }
                    else if (!SameFileStem(targetName, partName) &&
                             !SameFileStem(targetName, copyName))
                    {
                        row.ValidationStatus = "REVIEW";
                        row.ValidationMessage = Append(row.ValidationMessage,
                            "The target does not match either Name of the Part or Copy Version Name.");
                    }
                    else if (SameFileStem(targetName, partName) &&
                             SameFileStem(targetName, copyName))
                    {
                        row.ValidationStatus = "REVIEW";
                        row.ValidationMessage = Append(row.ValidationMessage,
                            "Original and copied names resolve to the same name.");
                    }

                    // Treat the last column as the desired replacement. The
                    // other member of the pair is the current/source reference.
                    if (SameFileStem(targetName, partName))
                    {
                        row.SourcePart = copyName;
                        row.ReplacementPart = partName;
                    }
                    else if (SameFileStem(targetName, copyName))
                    {
                        row.SourcePart = partName;
                        row.ReplacementPart = copyName;
                    }

                    rows.Add(row);
                }

                return rows;
            }
        }

        private static XDocument LoadXml(ZipArchive archive, string entryName)
        {
            ZipArchiveEntry entry = archive.GetEntry(entryName);
            if (entry == null)
                throw new InvalidDataException("Workbook component missing: " + entryName);
            using (Stream stream = entry.Open())
                return XDocument.Load(stream);
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return new List<string>();

            using (Stream stream = entry.Open())
            {
                XDocument doc = XDocument.Load(stream);
                return doc.Descendants(MainNs + "si")
                    .Select(si => String.Concat(
                        si.Descendants(MainNs + "t").Select(t => (string)t)))
                    .ToList();
            }
        }

        private static List<List<string>> ReadGrid(
            XDocument sheet, List<string> sharedStrings)
        {
            List<List<string>> result = new List<List<string>>();
            foreach (XElement row in sheet.Descendants(MainNs + "row"))
            {
                Dictionary<int, string> values = new Dictionary<int, string>();
                int maxCol = -1;
                foreach (XElement cell in row.Elements(MainNs + "c"))
                {
                    string address = (string)cell.Attribute("r") ?? "";
                    int col = ColumnIndex(address);
                    if (col < 0) continue;

                    string type = (string)cell.Attribute("t") ?? "";
                    string value = (string)cell.Element(MainNs + "v") ?? "";
                    if (type == "s")
                    {
                        int index;
                        value = Int32.TryParse(value, out index) &&
                                index >= 0 && index < sharedStrings.Count
                            ? sharedStrings[index] : "";
                    }
                    else if (type == "inlineStr")
                    {
                        value = String.Concat(
                            cell.Descendants(MainNs + "t").Select(t => (string)t));
                    }
                    else if (type == "b")
                    {
                        value = value == "1" ? "TRUE" : "FALSE";
                    }

                    values[col] = value;
                    if (col > maxCol) maxCol = col;
                }

                List<string> cells = Enumerable.Repeat("", maxCol + 1).ToList();
                foreach (KeyValuePair<int, string> pair in values)
                    cells[pair.Key] = pair.Value;
                result.Add(cells);
            }
            return result;
        }

        private static int ColumnIndex(string cellAddress)
        {
            int index = 0;
            int letters = 0;
            foreach (char c in cellAddress)
            {
                if (c < 'A' || c > 'Z')
                {
                    if (c >= 'a' && c <= 'z')
                        index = index * 26 + (Char.ToUpperInvariant(c) - 'A' + 1);
                    else
                        break;
                    letters++;
                }
                else
                {
                    index = index * 26 + (c - 'A' + 1);
                    letters++;
                }
            }
            return letters == 0 ? -1 : index - 1;
        }

        private static int FindHeader(List<string> headers, string name)
        {
            for (int i = 0; i < headers.Count; i++)
                if (String.Equals(headers[i].Trim().TrimStart('\uFEFF'),
                    name, StringComparison.OrdinalIgnoreCase))
                    return i;
            return -1;
        }

        private static string Cell(List<string> cells, int index)
        {
            return index >= 0 && index < cells.Count ? cells[index] ?? "" : "";
        }

        private static bool SameFileStem(string a, string b)
        {
            return String.Equals(NormalizePartName(a), NormalizePartName(b),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizePartName(string value)
        {
            value = (value ?? "").Trim();
            if (value.Length == 0) return "";
            return Path.GetFileNameWithoutExtension(value);
        }

        private static string NormalizeWorksheetPath(string target)
        {
            target = (target ?? "").Replace('\\', '/');
            if (target.StartsWith("/")) return target.TrimStart('/');
            if (target.StartsWith("xl/")) return target;
            return "xl/" + target.TrimStart('/');
        }

        private static string Append(string existing, string next)
        {
            if (String.IsNullOrWhiteSpace(existing)) return next;
            return existing + " " + next;
        }
    }
}
