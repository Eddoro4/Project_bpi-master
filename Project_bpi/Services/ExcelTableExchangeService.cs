using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Project_bpi.Services
{
    public sealed class ExcelTableData
    {
        public int ColumnCount { get; set; }
        public int HeaderRowCount { get; set; }
        public int BodyRowCount { get; set; }
        public List<ExcelTableCell> HeaderCells { get; } = new List<ExcelTableCell>();
        public List<ExcelTableCell> BodyCells { get; } = new List<ExcelTableCell>();
    }

    public sealed class ExcelTableCell
    {
        public string Text { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        public int ColSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
        public bool IsHeader { get; set; }
    }

    public static class ExcelTableExchangeService
    {
        private const string PackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        private const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        private const string SpreadsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";
        private const string XmlContentType = "application/xml";

        private sealed class WorksheetCellRecord
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public string Text { get; set; }
            public int StyleIndex { get; set; }
            public int ColSpan { get; set; } = 1;
            public int RowSpan { get; set; } = 1;
        }

        public static void Export(string outputPath, ExcelTableData tableData)
        {
            if (tableData == null)
            {
                throw new ArgumentNullException(nameof(tableData));
            }

            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var worksheetRows = BuildWorksheetRows(tableData);
            var sharedStrings = CollectSharedStrings(worksheetRows);

            using (var package = Package.Open(outputPath, FileMode.Create, FileAccess.ReadWrite))
            {
                var workbookUri = PackUriHelper.CreatePartUri(new Uri("/xl/workbook.xml", UriKind.Relative));
                var workbookPart = package.CreatePart(workbookUri, SpreadsheetContentType, CompressionOption.Maximum);
                WritePart(workbookPart, BuildWorkbookXml());

                var worksheetUri = PackUriHelper.CreatePartUri(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative));
                var worksheetPart = package.CreatePart(worksheetUri, WorksheetContentType, CompressionOption.Maximum);
                WritePart(worksheetPart, BuildWorksheetXml(worksheetRows, tableData.ColumnCount));

                var stylesUri = PackUriHelper.CreatePartUri(new Uri("/xl/styles.xml", UriKind.Relative));
                var stylesPart = package.CreatePart(stylesUri, StylesContentType, CompressionOption.Maximum);
                WritePart(stylesPart, BuildStylesXml());

                var sharedStringsUri = PackUriHelper.CreatePartUri(new Uri("/xl/sharedStrings.xml", UriKind.Relative));
                var sharedStringsPart = package.CreatePart(sharedStringsUri, SharedStringsContentType, CompressionOption.Maximum);
                WritePart(sharedStringsPart, BuildSharedStringsXml(sharedStrings));

                package.CreateRelationship(workbookUri, TargetMode.Internal, PackageRelationshipType);
                workbookPart.CreateRelationship(
                    PackUriHelper.GetRelativeUri(workbookPart.Uri, worksheetUri),
                    TargetMode.Internal,
                    WorksheetRelationshipType,
                    "rId1");
                workbookPart.CreateRelationship(
                    PackUriHelper.GetRelativeUri(workbookPart.Uri, stylesUri),
                    TargetMode.Internal,
                    StylesRelationshipType,
                    "rId2");
                workbookPart.CreateRelationship(
                    PackUriHelper.GetRelativeUri(workbookPart.Uri, sharedStringsUri),
                    TargetMode.Internal,
                    SharedStringsRelationshipType,
                    "rId3");
            }
        }

        public static ExcelTableData Import(string inputPath)
        {
            using (var package = Package.Open(inputPath, FileMode.Open, FileAccess.Read))
            {
                var workbookRelationship = package.GetRelationshipsByType(PackageRelationshipType).FirstOrDefault();
                if (workbookRelationship == null)
                {
                    throw new InvalidOperationException("Файл Excel не содержит книгу.");
                }

                var workbookUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), workbookRelationship.TargetUri);
                var workbookPart = package.GetPart(workbookUri);

                var workbookRelationships = workbookPart.GetRelationships()
                    .Select(item => (item.RelationshipType, PackUriHelper.ResolvePartUri(workbookPart.Uri, item.TargetUri)))
                    .ToList();
                var worksheetRelationship = workbookRelationships.FirstOrDefault(item => item.RelationshipType == WorksheetRelationshipType);
                if (worksheetRelationship.Item2 == null)
                {
                    throw new InvalidOperationException("Файл Excel не содержит лист с данными.");
                }

                var worksheetPart = package.GetPart(worksheetRelationship.Item2);
                var sharedStringsPart = workbookRelationships
                    .Where(item => item.RelationshipType == SharedStringsRelationshipType)
                    .Select(item => package.GetPart(item.Item2))
                    .FirstOrDefault();

                List<string> sharedStrings = sharedStringsPart != null
                    ? LoadSharedStrings(sharedStringsPart)
                    : new List<string>();

                return ParseWorksheet(worksheetPart, sharedStrings);
            }
        }

        private static List<List<WorksheetCellRecord>> BuildWorksheetRows(ExcelTableData tableData)
        {
            var rows = new List<List<WorksheetCellRecord>>();
            int columnCount = Math.Max(1, tableData.ColumnCount);

            for (int row = 1; row <= tableData.HeaderRowCount; row++)
            {
                rows.Add(BuildWorksheetRow(
                    tableData.HeaderCells.Where(cell => cell.Row == row).ToList(),
                    columnCount,
                    row,
                    0));
            }

            for (int row = 1; row <= tableData.BodyRowCount; row++)
            {
                rows.Add(BuildWorksheetRow(
                    tableData.BodyCells.Where(cell => cell.Row == row).ToList(),
                    columnCount,
                    tableData.HeaderRowCount + row,
                    0));
            }

            return rows;
        }

        private static List<WorksheetCellRecord> BuildWorksheetRow(
            List<ExcelTableCell> cells,
            int columnCount,
            int worksheetRow,
            int defaultStyleIndex)
        {
            var row = new List<WorksheetCellRecord>();

            foreach (var cell in cells.OrderBy(item => item.Column))
            {
                row.Add(new WorksheetCellRecord
                {
                    Row = worksheetRow,
                    Column = cell.Column,
                    Text = cell.Text ?? string.Empty,
                    StyleIndex = cell.IsHeader ? 1 : defaultStyleIndex,
                    ColSpan = Math.Max(1, cell.ColSpan),
                    RowSpan = Math.Max(1, cell.RowSpan)
                });
            }

            return row;
        }

        private static Dictionary<string, int> CollectSharedStrings(IEnumerable<List<WorksheetCellRecord>> rows)
        {
            var map = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (var cell in rows.SelectMany(item => item))
            {
                string text = cell.Text ?? string.Empty;
                if (!map.ContainsKey(text))
                {
                    map[text] = map.Count;
                }
            }

            return map;
        }

        private static string BuildWorkbookXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                   "<sheets><sheet name=\"Table\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
        }

        private static string BuildStylesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                   "<fonts count=\"2\">" +
                   "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
                   "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
                   "</fonts>" +
                   "<fills count=\"2\">" +
                   "<fill><patternFill patternType=\"none\"/></fill>" +
                   "<fill><patternFill patternType=\"gray125\"/></fill>" +
                   "</fills>" +
                   "<borders count=\"1\"><border><left style=\"thin\"/><right style=\"thin\"/><top style=\"thin\"/><bottom style=\"thin\"/><diagonal/></border></borders>" +
                   "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                   "<cellXfs count=\"1\">" +
                   "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>" +
                   "</cellXfs>" +
                   "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                   "</styleSheet>";
        }

        private static string BuildSharedStringsXml(Dictionary<string, int> sharedStrings)
        {
            var ordered = sharedStrings.OrderBy(item => item.Value).Select(item => item.Key).ToList();
            var builder = new StringBuilder();
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
                .Append(ordered.Count)
                .Append("\" uniqueCount=\"")
                .Append(ordered.Count)
                .Append("\">");

            foreach (string value in ordered)
            {
                builder.Append("<si><t xml:space=\"preserve\">")
                    .Append(Escape(value))
                    .Append("</t></si>");
            }

            builder.Append("</sst>");
            return builder.ToString();
        }

        private static string BuildWorksheetXml(List<List<WorksheetCellRecord>> rows, int columnCount)
        {
            var builder = new StringBuilder();
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            builder.Append("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>");
            builder.Append("<sheetFormatPr defaultRowHeight=\"18\"/>");
            builder.Append("<cols>");
            for (int column = 1; column <= Math.Max(1, columnCount); column++)
            {
                builder.Append("<col min=\"").Append(column).Append("\" max=\"").Append(column).Append("\" width=\"20\" customWidth=\"1\"/>");
            }
            builder.Append("</cols>");
            builder.Append("<sheetData>");

            foreach (var row in rows.Where(item => item.Any()))
            {
                int rowIndex = row.First().Row;
                builder.Append("<row r=\"").Append(rowIndex).Append("\" ht=\"22\" customHeight=\"1\">");

                foreach (var cell in row.OrderBy(item => item.Column))
                {
                    builder.Append("<c r=\"")
                        .Append(GetCellReference(cell.Column, cell.Row))
                        .Append("\" t=\"inlineStr\" s=\"")
                        .Append(cell.StyleIndex)
                        .Append("\"><is><t xml:space=\"preserve\">")
                        .Append(Escape(cell.Text))
                        .Append("</t></is></c>");
                }

                builder.Append("</row>");
            }

            builder.Append("</sheetData>");

            var mergedCells = rows
                .SelectMany(item => item)
                .Where(cell => cell.ColSpan > 1 || cell.RowSpan > 1)
                .ToList();

            if (mergedCells.Any())
            {
                builder.Append("<mergeCells count=\"").Append(mergedCells.Count).Append("\">");
                foreach (var cell in mergedCells)
                {
                    builder.Append("<mergeCell ref=\"")
                        .Append(GetCellReference(cell.Column, cell.Row))
                        .Append(":")
                        .Append(GetCellReference(cell.Column + cell.ColSpan - 1, cell.Row + cell.RowSpan - 1))
                        .Append("\"/>");
                }
                builder.Append("</mergeCells>");
            }

            builder.Append("</worksheet>");
            return builder.ToString();
        }

        private static List<string> LoadSharedStrings(PackagePart part)
        {
            var document = new XmlDocument();
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                document.Load(stream);
            }

            var manager = new XmlNamespaceManager(document.NameTable);
            manager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            return document.SelectNodes("//x:si", manager)
                .Cast<XmlNode>()
                .Select(GetInnerText)
                .ToList();
        }

        private static ExcelTableData ParseWorksheet(PackagePart worksheetPart, List<string> sharedStrings)
        {
            var document = new XmlDocument();
            using (var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                document.Load(stream);
            }

            var manager = new XmlNamespaceManager(document.NameTable);
            manager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var records = new List<WorksheetCellRecord>();
            foreach (XmlNode cellNode in document.SelectNodes("//x:sheetData/x:row/x:c", manager))
            {
                string cellReference = cellNode.Attributes?["r"]?.Value;
                if (string.IsNullOrWhiteSpace(cellReference))
                {
                    continue;
                }

                GetCoordinates(cellReference, out int row, out int column);

                string cellType = cellNode.Attributes?["t"]?.Value ?? string.Empty;
                int styleIndex = 0;
                int.TryParse(cellNode.Attributes?["s"]?.Value, out styleIndex);

                records.Add(new WorksheetCellRecord
                {
                    Row = row,
                    Column = column,
                    Text = ReadCellText(cellNode, manager, cellType, sharedStrings),
                    StyleIndex = styleIndex
                });
            }

            foreach (XmlNode mergeNode in document.SelectNodes("//x:mergeCells/x:mergeCell", manager))
            {
                string reference = mergeNode.Attributes?["ref"]?.Value;
                if (string.IsNullOrWhiteSpace(reference))
                {
                    continue;
                }

                string[] parts = reference.Split(':');
                if (parts.Length != 2)
                {
                    continue;
                }

                GetCoordinates(parts[0], out int startRow, out int startColumn);
                GetCoordinates(parts[1], out int endRow, out int endColumn);

                var topLeftCell = records.FirstOrDefault(item => item.Row == startRow && item.Column == startColumn);
                if (topLeftCell != null)
                {
                    topLeftCell.ColSpan = Math.Max(1, endColumn - startColumn + 1);
                    topLeftCell.RowSpan = Math.Max(1, endRow - startRow + 1);
                }
            }

            return ConvertRecordsToTableData(records);
        }

        private static ExcelTableData ConvertRecordsToTableData(List<WorksheetCellRecord> records)
        {
            var tableData = new ExcelTableData();
            if (!records.Any())
            {
                return tableData;
            }

            int maxColumn = records.Max(item => item.Column + item.ColSpan - 1);
            int maxRow = records.Max(item => item.Row + item.RowSpan - 1);
            int numberingRow = FindNumberingRow(records, maxColumn);
            int headerRowCount = numberingRow > 1 ? numberingRow - 1 : InferHeaderRowCount(records);
            if (headerRowCount <= 0 && maxRow > 0)
            {
                headerRowCount = 1;
            }

            tableData.ColumnCount = Math.Max(1, maxColumn);
            tableData.HeaderRowCount = Math.Max(1, headerRowCount);

            foreach (var record in records.OrderBy(item => item.Row).ThenBy(item => item.Column))
            {
                if (numberingRow > 0 && record.Row == numberingRow)
                {
                    continue;
                }

                bool isHeaderRow = record.Row <= headerRowCount;
                int adjustedRow = isHeaderRow
                    ? record.Row
                    : record.Row - headerRowCount - (numberingRow > 0 && record.Row > numberingRow ? 1 : 0);

                var cell = new ExcelTableCell
                {
                    Text = record.Text,
                    Column = record.Column,
                    Row = Math.Max(1, adjustedRow),
                    ColSpan = record.ColSpan,
                    RowSpan = record.RowSpan,
                    IsHeader = isHeaderRow
                };

                if (isHeaderRow)
                {
                    tableData.HeaderCells.Add(cell);
                }
                else
                {
                    tableData.BodyCells.Add(cell);
                }
            }

            tableData.BodyRowCount = tableData.BodyCells.Any()
                ? tableData.BodyCells.Max(item => item.Row + item.RowSpan - 1)
                : Math.Max(0, maxRow - tableData.HeaderRowCount - (numberingRow > 0 ? 1 : 0));

            return tableData;
        }

        private static int FindNumberingRow(List<WorksheetCellRecord> records, int columnCount)
        {
            foreach (var rowGroup in records.GroupBy(item => item.Row).OrderBy(item => item.Key))
            {
                var ordered = rowGroup.OrderBy(item => item.Column).ToList();
                if (ordered.Count < columnCount)
                {
                    continue;
                }

                bool isNumbering = true;
                for (int column = 1; column <= columnCount; column++)
                {
                    var cell = ordered.FirstOrDefault(item => item.Column == column);
                    if (cell == null || !string.Equals(cell.Text?.Trim(), column.ToString(), StringComparison.Ordinal))
                    {
                        isNumbering = false;
                        break;
                    }
                }

                if (isNumbering)
                {
                    return rowGroup.Key;
                }
            }

            return 0;
        }

        private static int InferHeaderRowCount(List<WorksheetCellRecord> records)
        {
            int headerRow = 0;
            foreach (var rowGroup in records.GroupBy(item => item.Row).OrderBy(item => item.Key))
            {
                if (rowGroup.All(item => item.StyleIndex == 1))
                {
                    headerRow = rowGroup.Key;
                    continue;
                }

                break;
            }

            return headerRow;
        }

        private static string ReadCellText(XmlNode cellNode, XmlNamespaceManager manager, string cellType, List<string> sharedStrings)
        {
            if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase))
            {
                string indexText = cellNode.SelectSingleNode("x:v", manager)?.InnerText ?? "0";
                return int.TryParse(indexText, out int index) && index >= 0 && index < sharedStrings.Count
                    ? sharedStrings[index]
                    : string.Empty;
            }

            if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                return GetInnerText(cellNode.SelectSingleNode("x:is", manager));
            }

            return cellNode.SelectSingleNode("x:v", manager)?.InnerText ?? string.Empty;
        }

        private static string GetInnerText(XmlNode node)
        {
            if (node == null)
            {
                return string.Empty;
            }

            return string.Concat(node.ChildNodes.Cast<XmlNode>().Select(child =>
                child.Name.EndsWith("t", StringComparison.OrdinalIgnoreCase)
                    ? child.InnerText
                    : GetInnerText(child)));
        }

        private static void GetCoordinates(string cellReference, out int row, out int column)
        {
            var match = Regex.Match(cellReference ?? string.Empty, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                row = 1;
                column = 1;
                return;
            }

            column = GetColumnNumber(match.Groups[1].Value);
            row = int.Parse(match.Groups[2].Value);
        }

        private static string GetCellReference(int column, int row)
        {
            return GetColumnName(column) + row;
        }

        private static int GetColumnNumber(string columnName)
        {
            int value = 0;
            foreach (char character in columnName.ToUpperInvariant())
            {
                value = value * 26 + (character - 'A' + 1);
            }

            return value;
        }

        private static string GetColumnName(int column)
        {
            var builder = new StringBuilder();
            int current = column;
            while (current > 0)
            {
                current--;
                builder.Insert(0, (char)('A' + current % 26));
                current /= 26;
            }

            return builder.ToString();
        }

        private static void WritePart(PackagePart part, string xml)
        {
            using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(xml);
            }
        }

        private static string Escape(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }
}
