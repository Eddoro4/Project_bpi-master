using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Project_bpi.Models;
using Project_bpi.Services;

namespace Project_bpi
{
    public partial class MainWindow
    {
        private static TableSeed CreateSimpleTableSeed(string title, string patternName, params string[] headers)
        {
            return new TableSeed
            {
                Title = title,
                PatternName = patternName,
                Cells = headers
                    .Select((header, index) => new TableCellSeed
                    {
                        Row = 1,
                        Column = index + 1,
                        Text = header,
                        ColSpan = 1,
                        RowSpan = 1,
                        IsHeader = true
                    })
                    .ToArray()
            };
        }

        private async Task EnsureNirPublishingStaffReportInDatabaseAsync()
        {
            var database = new DataBase(GetSharedTemplateDatabasePath());
            database.InitializeDatabase(false);

            if (await database.HasSuppressedBuiltInReport(NirPublishingStaffReportSuppressionKey))
            {
                return;
            }

            var reports = await database.GetAllReports();
            var existingReport = reports.FirstOrDefault(candidateReport =>
                string.Equals(candidateReport.Title, NirPublishingStaffReportTitle, StringComparison.Ordinal));

            var sourceSubSection = await TryLoadNirPublishingStaffSubSectionAsync(database, reports);

            int reportId;
            if (existingReport == null)
            {
                int patternId = await database.AddFilePattern(new PatternFile
                {
                    Title = NirPublishingStaffReportTitle,
                    Year = DateTime.Today.Year,
                    Path = database.DatabasePath
                });

                reportId = await database.AddReport(new Report
                {
                    Title = NirPublishingStaffReportTitle,
                    Year = DateTime.Today.Year,
                    PattarnId = patternId
                });
            }
            else
            {
                reportId = existingReport.Id;
            }

            var staffReport = await database.GetFullReport(reportId);
            if (staffReport == null)
            {
                return;
            }

            foreach (var section in staffReport.Sections
                .Where(section => section.Number != 1)
                .OrderByDescending(section => section.Number)
                .ToList())
            {
                await database.DeleteSection(section.Id);
            }

            staffReport = await database.GetFullReport(reportId);
            if (staffReport == null)
            {
                return;
            }

            var sectionSeed = CreateNirPublishingStaffReportSection();
            var targetSection = await EnsureNirSectionAsync(database, staffReport, sectionSeed);
            var targetContentSubSection = await EnsureSectionContentSubSectionAsync(database, targetSection);

            if (sourceSubSection != null && !HasStoredSubSectionContent(targetContentSubSection))
            {
                await ClearSubSectionContentAsync(database, targetContentSubSection);
                await CloneSubSectionContentAsync(database, sourceSubSection, targetContentSubSection);
                return;
            }

            await EnsureNirSectionContentAsync(database, targetSection, sectionSeed);
        }

        private async Task<Section> EnsureNirSectionAsync(DataBase database, Report report, NirSectionSeed seed)
        {
            var section = report.Sections.FirstOrDefault(item => item.Number == seed.Number);
            if (section == null)
            {
                int sectionId = await database.AddSection(new Section
                {
                    ReportId = report.Id,
                    Number = seed.Number,
                    Title = seed.Title
                });

                section = new Section
                {
                    Id = sectionId,
                    ReportId = report.Id,
                    Number = seed.Number,
                    Title = seed.Title,
                    Report = report,
                    SubSections = new List<SubSection>()
                };

                report.Sections.Add(section);
            }
            else if (ShouldUpdateSeededSectionTitle(section, seed.Title))
            {
                section.Title = seed.Title;
                await database.UpdateSection(section);
            }

            if (section.SubSections == null)
            {
                section.SubSections = new List<SubSection>();
            }

            return section;
        }

        private async Task EnsureNirSectionContentAsync(DataBase database, Section section, NirSectionSeed seed)
        {
            if (seed.Tables == null || seed.Tables.Length == 0)
            {
                return;
            }

            var contentSubSection = await EnsureSectionContentSubSectionAsync(database, section);
            foreach (var tableSeed in seed.Tables)
            {
                await EnsureSeededTableAsync(database, contentSubSection, tableSeed);
            }
        }

        private async Task<SubSection> EnsureSectionContentSubSectionAsync(DataBase database, Section section)
        {
            var existingSubSection = GetSectionContentSubSection(section);
            if (existingSubSection != null)
            {
                return existingSubSection;
            }

            int subsectionId = await database.AddSubsection(new SubSection
            {
                SectionId = section.Id,
                ParentSubsectionId = null,
                Number = 0,
                Title = SectionContentSubSectionTitle
            });

            var newSubSection = new SubSection
            {
                Id = subsectionId,
                SectionId = section.Id,
                ParentSubsectionId = null,
                Number = 0,
                Title = SectionContentSubSectionTitle,
                Section = section,
                SubSections = new List<SubSection>(),
                Tables = new List<Table>(),
                Texts = new List<Text>()
            };

            section.SubSections.Add(newSubSection);
            return newSubSection;
        }

        private async Task EnsureSeededTableAsync(DataBase database, SubSection subsection, TableSeed seed)
        {
            if (seed == null || string.IsNullOrWhiteSpace(seed.PatternName))
            {
                return;
            }

            if (subsection.Tables == null)
            {
                subsection.Tables = new List<Table>();
            }

            var existingTable = subsection.Tables.FirstOrDefault(item =>
                string.Equals(item.PatternName, seed.PatternName, StringComparison.Ordinal));

            if (existingTable == null)
            {
                int tableId = await database.AddTable(new Table
                {
                    Title = seed.Title,
                    SubsectionId = subsection.Id,
                    PatternName = seed.PatternName
                });

                existingTable = new Table
                {
                    Id = tableId,
                    Title = seed.Title,
                    SubsectionId = subsection.Id,
                    PatternName = seed.PatternName,
                    Subsection = subsection,
                    TableItems = new List<TableItem>()
                };

                subsection.Tables.Add(existingTable);
            }

            if (existingTable.TableItems == null)
            {
                existingTable.TableItems = new List<TableItem>();
            }

            if (existingTable.TableItems.Any() || seed.Cells == null || seed.Cells.Length == 0)
            {
                return;
            }

            foreach (var cell in seed.Cells)
            {
                int itemId = await database.AddTableItem(new TableItem
                {
                    TableId = existingTable.Id,
                    Row = cell.Row,
                    Column = cell.Column,
                    Header = cell.Text ?? string.Empty,
                    ColSpan = cell.ColSpan,
                    RowSpan = cell.RowSpan,
                    IsHeader = cell.IsHeader
                });

                existingTable.TableItems.Add(new TableItem
                {
                    Id = itemId,
                    TableId = existingTable.Id,
                    Row = cell.Row,
                    Column = cell.Column,
                    Header = cell.Text ?? string.Empty,
                    ColSpan = cell.ColSpan,
                    RowSpan = cell.RowSpan,
                    IsHeader = cell.IsHeader,
                    Table = existingTable
                });
            }
        }

        private async Task<SubSection> TryLoadNirPublishingStaffSubSectionAsync(
            DataBase database,
            IEnumerable<Report> reports)
        {
            var nirReportInfo = reports.FirstOrDefault(report =>
                string.Equals(report.Title, NirReportTitle, StringComparison.Ordinal));
            if (nirReportInfo == null)
            {
                return null;
            }

            var nirReport = await database.GetFullReport(nirReportInfo.Id);
            return FindNirPublishingStaffSubSection(nirReport);
        }

        private SubSection FindNirPublishingStaffSubSection(Report report)
        {
            var publishingSection = report?.Sections?.FirstOrDefault(section => section.Number == 3);
            return publishingSection?.SubSections?.FirstOrDefault(subsection =>
                !IsSectionContentSubSection(subsection) &&
                !subsection.ParentSubsectionId.HasValue &&
                string.Equals(subsection.Title, NirPublishingStaffReportTitle, StringComparison.Ordinal));
        }

        private bool HasStoredSubSectionContent(SubSection subsection)
        {
            if (subsection == null)
            {
                return false;
            }

            bool hasTexts = subsection.Texts != null && subsection.Texts.Any(text =>
                !string.IsNullOrWhiteSpace(text?.Content));
            bool hasTables = subsection.Tables != null && subsection.Tables.Any(table =>
                table?.TableItems != null && table.TableItems.Any(item => !item.IsHeader));
            return hasTexts || hasTables;
        }

        private async Task CloneSubSectionContentAsync(DataBase database, SubSection source, SubSection target)
        {
            if (database == null || source == null || target == null)
            {
                return;
            }

            if (target.Texts == null)
            {
                target.Texts = new List<Text>();
            }

            if (target.Tables == null)
            {
                target.Tables = new List<Table>();
            }

            foreach (var text in source.Texts ?? Enumerable.Empty<Text>())
            {
                int textId = await database.AddText(new Text
                {
                    SubsectionId = target.Id,
                    Content = text.Content,
                    PatternName = text.PatternName
                });

                target.Texts.Add(new Text
                {
                    Id = textId,
                    SubsectionId = target.Id,
                    Content = text.Content,
                    PatternName = text.PatternName,
                    Subsection = target
                });
            }

            foreach (var table in source.Tables ?? Enumerable.Empty<Table>())
            {
                int tableId = await database.AddTable(new Table
                {
                    Title = table.Title,
                    SubsectionId = target.Id,
                    PatternName = table.PatternName
                });

                var clonedTable = new Table
                {
                    Id = tableId,
                    Title = table.Title,
                    SubsectionId = target.Id,
                    PatternName = table.PatternName,
                    Subsection = target,
                    TableItems = new List<TableItem>()
                };

                target.Tables.Add(clonedTable);

                foreach (var item in table.TableItems ?? Enumerable.Empty<TableItem>())
                {
                    int itemId = await database.AddTableItem(new TableItem
                    {
                        TableId = tableId,
                        Column = item.Column,
                        Row = item.Row,
                        Header = item.Header,
                        ColSpan = item.ColSpan,
                        RowSpan = item.RowSpan,
                        IsHeader = item.IsHeader
                    });

                    clonedTable.TableItems.Add(new TableItem
                    {
                        Id = itemId,
                        TableId = tableId,
                        Column = item.Column,
                        Row = item.Row,
                        Header = item.Header,
                        ColSpan = item.ColSpan,
                        RowSpan = item.RowSpan,
                        IsHeader = item.IsHeader,
                        Table = clonedTable
                    });
                }
            }
        }

        private async Task ClearSubSectionContentAsync(DataBase database, SubSection subsection)
        {
            if (database == null || subsection == null)
            {
                return;
            }

            foreach (var text in subsection.Texts?.ToList() ?? new List<Text>())
            {
                await database.DeleteText(text.Id);
            }

            foreach (var table in subsection.Tables?.ToList() ?? new List<Table>())
            {
                await database.DeleteTable(table.Id);
            }

            subsection.Texts = new List<Text>();
            subsection.Tables = new List<Table>();
        }

        private bool ShouldUpdateSeededSectionTitle(Section section, string seededTitle)
        {
            if (section == null || string.IsNullOrWhiteSpace(seededTitle))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(section.Title))
            {
                return true;
            }

            string currentTitle = section.Title.Trim();
            return string.Equals(currentTitle, GetDefaultSectionTitle(section.Number), StringComparison.Ordinal)
                && !string.Equals(currentTitle, seededTitle, StringComparison.Ordinal);
        }

        private static NirSectionSeed CreateNirPublishingStaffReportSection() => new NirSectionSeed
        {
            Number = 1,
            Title = NirPublishingStaffReportTitle,
            Tables = new[]
            {
                CreateSimpleTableSeed(
                    NirPublishingStaffReportTitle,
                    NirPublishingTablePatternName,
                    "№",
                    "Доля авторов заполняющей кафедры",
                    "Фамилия И.О. авторов",
                    "Наименование публикации",
                    "Тип публикации",
                    "Наименование издания, год, страницы публикации (без кавычек)",
                    "Место издания (наименование организации, город)")
            }
        };
    }
}
