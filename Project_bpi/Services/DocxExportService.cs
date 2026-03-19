using Project_bpi.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security;
using System.Text;

namespace Project_bpi.Services
{
    public static class DocxExportService
    {
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        private sealed class ExportCellSpec
        {
            public int Column { get; set; }
            public int ColSpan { get; set; } = 1;
            public string Text { get; set; }
            public bool IsHeader { get; set; }
            public bool IsVerticalRestart { get; set; }
            public bool IsVerticalContinuation { get; set; }
            public bool IsNumberingCell { get; set; }
        }

        public static void ExportReport(Report report, string outputPath)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("Не указан путь для сохранения документа.", nameof(outputPath));
            }

            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string documentXml = BuildDocumentXml(report);

            using (var package = Package.Open(outputPath, FileMode.Create, FileAccess.ReadWrite))
            {
                var documentUri = PackUriHelper.CreatePartUri(new Uri("/word/document.xml", UriKind.Relative));
                var documentPart = package.CreatePart(
                    documentUri,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    CompressionOption.Maximum);

                using (var stream = documentPart.GetStream(FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(documentXml);
                }

                package.CreateRelationship(documentUri, TargetMode.Internal, OfficeDocumentRelationshipType);
            }
        }

        private static string BuildDocumentXml(Report report)
        {
            var builder = new StringBuilder();

            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<w:document xmlns:w=\"").Append(WordNamespace).Append("\">");
            builder.Append("<w:body>");

            AppendParagraph(builder, report.Title, true, 32, true);

            foreach (var section in report.Sections?.OrderBy(item => item.Number).ThenBy(item => item.Id) ?? Enumerable.Empty<Section>())
            {
                AppendSection(builder, section);
            }

            builder.Append("<w:sectPr>");
            builder.Append("<w:pgSz w:w=\"11906\" w:h=\"16838\"/>");
            builder.Append("<w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"850\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>");
            builder.Append("</w:sectPr>");
            builder.Append("</w:body>");
            builder.Append("</w:document>");

            return builder.ToString();
        }

        private static void AppendSection(StringBuilder builder, Section section)
        {
            string sectionTitle = BuildSectionTitle(section);
            if (!string.IsNullOrWhiteSpace(sectionTitle))
            {
                AppendParagraph(builder, sectionTitle, true, 28, false);
            }

            var sectionContent = section.SubSections?.FirstOrDefault(item =>
                string.Equals(item.Title, "__section_content__", StringComparison.Ordinal));

            AppendSubSectionContent(builder, sectionContent, includeTitle: false);

            foreach (var subsection in section.SubSections?
                .Where(item => !string.Equals(item.Title, "__section_content__", StringComparison.Ordinal))
                .OrderBy(item => item.Number)
                .ThenBy(item => item.Id) ?? Enumerable.Empty<SubSection>())
            {
                AppendSubSection(builder, subsection, 1);
            }
        }

        private static void AppendSubSection(StringBuilder builder, SubSection subsection, int depth)
        {
            if (!string.IsNullOrWhiteSpace(subsection?.Title))
            {
                int size = depth > 1 ? 24 : 26;
                AppendParagraph(builder, subsection.Title.Trim(), true, size, false);
            }

            AppendSubSectionContent(builder, subsection, includeTitle: false);

            foreach (var child in subsection.SubSections?.OrderBy(item => item.Number).ThenBy(item => item.Id) ?? Enumerable.Empty<SubSection>())
            {
                AppendSubSection(builder, child, depth + 1);
            }
        }

        private static void AppendSubSectionContent(StringBuilder builder, SubSection subsection, bool includeTitle)
        {
            if (subsection == null)
            {
                return;
            }

            if (includeTitle && !string.IsNullOrWhiteSpace(subsection.Title))
            {
                AppendParagraph(builder, subsection.Title.Trim(), true, 26, false);
            }

            foreach (var text in subsection.Texts?.OrderBy(item => item.Id) ?? Enumerable.Empty<Text>())
            {
                AppendMultilineText(builder, text.Content);
            }

            foreach (var table in subsection.Tables?.OrderBy(item => item.Id) ?? Enumerable.Empty<Table>())
            {
                AppendTable(builder, table);
            }
        }

        private static void AppendMultilineText(StringBuilder builder, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            foreach (var line in NormalizeNewLines(text).Split('\n'))
            {
                AppendParagraph(builder, line, false, 24, false);
            }
        }

        private static void AppendTable(StringBuilder builder, Table table)
        {
            if (!string.IsNullOrWhiteSpace(table?.Title))
            {
                AppendParagraph(builder, table.Title.Trim(), true, 24, false);
            }

            var rows = BuildTableRows(table, out int columnCount, out int headerRowCount);
            if (rows.Count == 0 || columnCount <= 0)
            {
                return;
            }

            builder.Append("<w:tbl>");
            builder.Append("<w:tblPr>");
            builder.Append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
            builder.Append("<w:tblBorders>");
            builder.Append("<w:top w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>");
            builder.Append("<w:left w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>");
            builder.Append("<w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>");
            builder.Append("<w:right w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>");
            builder.Append("<w:insideH w:val=\"single\" w:sz=\"6\" w:space=\"0\" w:color=\"808080\"/>");
            builder.Append("<w:insideV w:val=\"single\" w:sz=\"6\" w:space=\"0\" w:color=\"808080\"/>");
            builder.Append("</w:tblBorders>");
            builder.Append("</w:tblPr>");

            builder.Append("<w:tblGrid>");
            for (int i = 0; i < columnCount; i++)
            {
                builder.Append("<w:gridCol w:w=\"2400\"/>");
            }
            builder.Append("</w:tblGrid>");

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                AppendTableRow(builder, rows[rowIndex]);

                if (headerRowCount > 0 && rowIndex == headerRowCount - 1)
                {
                    AppendTableRow(builder, BuildNumberingRow(columnCount));
                }
            }

            if (headerRowCount == 0)
            {
                AppendTableRow(builder, BuildNumberingRow(columnCount));
            }

            builder.Append("</w:tbl>");
            AppendParagraph(builder, string.Empty, false, 20, false);
        }

        private static List<List<ExportCellSpec>> BuildTableRows(Table table, out int columnCount, out int headerRowCount)
        {
            columnCount = 0;
            headerRowCount = 0;

            var items = table?.TableItems?.ToList() ?? new List<TableItem>();
            if (items.Count == 0)
            {
                return new List<List<ExportCellSpec>>();
            }

            int maxRow = 0;
            var rowMap = new Dictionary<int, List<ExportCellSpec>>();

            foreach (var item in items)
            {
                maxRow = Math.Max(maxRow, item.Row + Math.Max(item.RowSpan, 1) - 1);
                columnCount = Math.Max(columnCount, item.Column + Math.Max(item.ColSpan, 1) - 1);

                AddRowSpec(rowMap, item.Row, new ExportCellSpec
                {
                    Column = item.Column,
                    ColSpan = Math.Max(item.ColSpan, 1),
                    Text = item.Header,
                    IsHeader = item.IsHeader,
                    IsVerticalRestart = item.RowSpan > 1
                });

                if (item.IsHeader)
                {
                    headerRowCount = Math.Max(headerRowCount, item.Row + Math.Max(item.RowSpan, 1) - 1);
                }

                for (int extraRow = 1; extraRow < Math.Max(item.RowSpan, 1); extraRow++)
                {
                    AddRowSpec(rowMap, item.Row + extraRow, new ExportCellSpec
                    {
                        Column = item.Column,
                        ColSpan = Math.Max(item.ColSpan, 1),
                        Text = string.Empty,
                        IsHeader = item.IsHeader,
                        IsVerticalContinuation = true
                    });
                }
            }

            var rows = new List<List<ExportCellSpec>>();
            for (int row = 1; row <= maxRow; row++)
            {
                var rowCells = rowMap.ContainsKey(row)
                    ? rowMap[row].OrderBy(item => item.Column).ToList()
                    : new List<ExportCellSpec>();

                var normalizedRow = new List<ExportCellSpec>();
                int expectedColumn = 1;
                foreach (var cell in rowCells)
                {
                    while (expectedColumn < cell.Column)
                    {
                        normalizedRow.Add(new ExportCellSpec
                        {
                            Column = expectedColumn,
                            ColSpan = 1,
                            Text = string.Empty,
                            IsHeader = row <= headerRowCount
                        });
                        expectedColumn++;
                    }

                    normalizedRow.Add(cell);
                    expectedColumn = cell.Column + Math.Max(cell.ColSpan, 1);
                }

                while (expectedColumn <= columnCount)
                {
                    normalizedRow.Add(new ExportCellSpec
                    {
                        Column = expectedColumn,
                        ColSpan = 1,
                        Text = string.Empty,
                        IsHeader = row <= headerRowCount
                    });
                    expectedColumn++;
                }

                rows.Add(normalizedRow);
            }

            return rows;
        }

        private static void AddRowSpec(Dictionary<int, List<ExportCellSpec>> rowMap, int row, ExportCellSpec cell)
        {
            if (!rowMap.TryGetValue(row, out var list))
            {
                list = new List<ExportCellSpec>();
                rowMap[row] = list;
            }

            list.Add(cell);
        }

        private static List<ExportCellSpec> BuildNumberingRow(int columnCount)
        {
            var row = new List<ExportCellSpec>();

            for (int column = 1; column <= columnCount; column++)
            {
                row.Add(new ExportCellSpec
                {
                    Column = column,
                    ColSpan = 1,
                    Text = column.ToString(),
                    IsNumberingCell = true
                });
            }

            return row;
        }

        private static void AppendTableRow(StringBuilder builder, List<ExportCellSpec> row)
        {
            builder.Append("<w:tr>");

            foreach (var cell in row)
            {
                builder.Append("<w:tc><w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/>");

                if (cell.ColSpan > 1)
                {
                    builder.Append("<w:gridSpan w:val=\"").Append(cell.ColSpan).Append("\"/>");
                }

                if (cell.IsVerticalRestart)
                {
                    builder.Append("<w:vMerge w:val=\"restart\"/>");
                }
                else if (cell.IsVerticalContinuation)
                {
                    builder.Append("<w:vMerge/>");
                }

                if (cell.IsHeader)
                {
                    builder.Append("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"DDEBF7\"/>");
                }
                else if (cell.IsNumberingCell)
                {
                    builder.Append("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"F2F2F2\"/>");
                }

                builder.Append("</w:tcPr>");
                builder.Append("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r>");
                if (cell.IsHeader || cell.IsNumberingCell)
                {
                    builder.Append("<w:rPr><w:b/></w:rPr>");
                }

                builder.Append("<w:t xml:space=\"preserve\">")
                    .Append(Escape(cell.Text))
                    .Append("</w:t></w:r></w:p></w:tc>");
            }

            builder.Append("</w:tr>");
        }

        private static void AppendParagraph(StringBuilder builder, string text, bool bold, int halfPoints, bool center)
        {
            builder.Append("<w:p><w:pPr>");
            if (center)
            {
                builder.Append("<w:jc w:val=\"center\"/>");
            }

            builder.Append("<w:spacing w:after=\"120\"/>");
            builder.Append("</w:pPr><w:r><w:rPr>");

            if (bold)
            {
                builder.Append("<w:b/>");
            }

            builder.Append("<w:sz w:val=\"").Append(halfPoints).Append("\"/>");
            builder.Append("</w:rPr><w:t xml:space=\"preserve\">")
                .Append(Escape(text))
                .Append("</w:t></w:r></w:p>");
        }

        private static string BuildSectionTitle(Section section)
        {
            if (section == null)
            {
                return string.Empty;
            }

            if (section.Number <= 0)
            {
                return section.Title?.Trim() ?? string.Empty;
            }

            return string.IsNullOrWhiteSpace(section.Title)
                ? $"Раздел {section.Number}"
                : $"{section.Number} {section.Title.Trim()}";
        }

        private static string NormalizeNewLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
        }

        private static string Escape(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }
}
