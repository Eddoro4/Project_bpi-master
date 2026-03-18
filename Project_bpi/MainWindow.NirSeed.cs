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
        private static TableCellSeed HeaderCell(int row, int column, string text, int colSpan = 1, int rowSpan = 1)
        {
            return new TableCellSeed
            {
                Row = row,
                Column = column,
                Text = text,
                ColSpan = colSpan,
                RowSpan = rowSpan,
                IsHeader = true
            };
        }

        private static TableCellSeed BodyCell(int row, int column, string text, int colSpan = 1, int rowSpan = 1)
        {
            return new TableCellSeed
            {
                Row = row,
                Column = column,
                Text = text,
                ColSpan = colSpan,
                RowSpan = rowSpan,
                IsHeader = false
            };
        }

        private static TableSeed CreateSimpleTableSeed(string title, string patternName, params string[] headers)
        {
            return new TableSeed
            {
                Title = title,
                PatternName = patternName,
                Cells = headers
                    .Select((header, index) => HeaderCell(1, index + 1, header))
                    .ToArray()
            };
        }

        private void HideStaticNirReportMenu()
        {
            NIR_Header.Visibility = System.Windows.Visibility.Collapsed;
            NIR_Menu.Visibility = System.Windows.Visibility.Collapsed;
        }

        private async Task EnsureNirReportInDatabaseAsync()
        {
            var database = new DataBase(GetSharedTemplateDatabasePath());
            database.InitializeDatabase(false);

            var reports = await database.GetAllReports();
            var existingNirReport = reports.FirstOrDefault(report =>
                string.Equals(report.Title, NirReportTitle, StringComparison.Ordinal));

            int reportId;
            if (existingNirReport == null)
            {
                int patternId = await database.AddFilePattern(new PatternFile
                {
                    Title = NirReportTitle,
                    Year = DateTime.Today.Year,
                    Path = database.DatabasePath
                });

                reportId = await database.AddReport(new Report
                {
                    Title = NirReportTitle,
                    Year = DateTime.Today.Year,
                    PattarnId = patternId
                });
            }
            else
            {
                reportId = existingNirReport.Id;
            }

            var nirReport = await database.GetFullReport(reportId);
            if (nirReport == null)
            {
                return;
            }

            foreach (var sectionSeed in GetNirReportSections())
            {
                var section = await EnsureNirSectionAsync(database, nirReport, sectionSeed);
                await EnsureNirSectionContentAsync(database, section, sectionSeed);

                foreach (var subSectionSeed in sectionSeed.SubSections ?? Array.Empty<NirSubSectionSeed>())
                {
                    await EnsureNirSubSectionTreeAsync(database, section, null, subSectionSeed);
                }
            }
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
            if (string.IsNullOrWhiteSpace(seed.TextContent) && (seed.Tables == null || seed.Tables.Length == 0))
            {
                return;
            }

            var contentSubSection = await EnsureSectionContentSubSectionAsync(database, section);
            await EnsureSeededTextAsync(database, contentSubSection, seed.TextPatternName, seed.TextContent);

            foreach (var tableSeed in seed.Tables ?? Array.Empty<TableSeed>())
            {
                await EnsureSeededTableAsync(database, contentSubSection, tableSeed);
            }
        }

        private async Task EnsureNirSubSectionTreeAsync(DataBase database, Section section, SubSection parent, NirSubSectionSeed seed)
        {
            var siblings = parent == null ? section.SubSections : parent.SubSections;
            if (siblings == null)
            {
                siblings = new List<SubSection>();
                if (parent == null)
                {
                    section.SubSections = siblings;
                }
                else
                {
                    parent.SubSections = siblings;
                }
            }

            var subsection = siblings.FirstOrDefault(item =>
                !IsSectionContentSubSection(item) &&
                item.ParentSubsectionId == parent?.Id &&
                (item.Number == seed.Number || string.Equals(item.Title, seed.Title, StringComparison.Ordinal)));

            if (subsection == null)
            {
                int subsectionId = await database.AddSubsection(new SubSection
                {
                    SectionId = section.Id,
                    ParentSubsectionId = parent?.Id,
                    Number = seed.Number,
                    Title = seed.Title
                });

                subsection = new SubSection
                {
                    Id = subsectionId,
                    SectionId = section.Id,
                    ParentSubsectionId = parent?.Id,
                    Number = seed.Number,
                    Title = seed.Title,
                    Section = section,
                    ParentSubsection = parent,
                    SubSections = new List<SubSection>(),
                    Tables = new List<Table>(),
                    Texts = new List<Text>()
                };

                siblings.Add(subsection);
            }
            else if (string.IsNullOrWhiteSpace(subsection.Title))
            {
                subsection.Title = seed.Title;
                await database.UpdateSubsection(subsection);
            }

            if (subsection.SubSections == null)
            {
                subsection.SubSections = new List<SubSection>();
            }

            if (subsection.Tables == null)
            {
                subsection.Tables = new List<Table>();
            }

            if (subsection.Texts == null)
            {
                subsection.Texts = new List<Text>();
            }

            await EnsureSeededTextAsync(database, subsection, seed.TextPatternName, seed.TextContent);

            foreach (var tableSeed in seed.Tables ?? Array.Empty<TableSeed>())
            {
                await EnsureSeededTableAsync(database, subsection, tableSeed);
            }

            foreach (var childSeed in seed.Children ?? Array.Empty<NirSubSectionSeed>())
            {
                await EnsureNirSubSectionTreeAsync(database, section, subsection, childSeed);
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

        private async Task EnsureSeededTextAsync(DataBase database, SubSection subsection, string patternName, string content)
        {
            if (string.IsNullOrWhiteSpace(patternName) || string.IsNullOrWhiteSpace(content))
            {
                return;
            }

            var existingText = subsection.Texts?.FirstOrDefault(item =>
                string.Equals(item.PatternName, patternName, StringComparison.Ordinal))
                ?? subsection.Texts?.FirstOrDefault(item => !string.IsNullOrWhiteSpace(item.Content));

            if (existingText != null)
            {
                return;
            }

            int textId = await database.AddText(new Text
            {
                SubsectionId = subsection.Id,
                Content = content,
                PatternName = patternName
            });

            subsection.Texts.Add(new Text
            {
                Id = textId,
                SubsectionId = subsection.Id,
                Content = content,
                PatternName = patternName,
                Subsection = subsection
            });
        }

        private async Task EnsureSeededTableAsync(DataBase database, SubSection subsection, TableSeed seed)
        {
            if (seed == null || string.IsNullOrWhiteSpace(seed.PatternName))
            {
                return;
            }

            var existingTable = subsection.Tables?.FirstOrDefault(item =>
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

        private static NirSectionSeed[] GetNirReportSections()
        {
            return new[]
            {
                CreateNirTitleSection(),
                CreateNirSection1(),
                CreateNirSection2(),
                CreateNirSection3(),
                CreateNirSection4(),
                CreateNirSection5(),
                CreateNirSection6(),
                CreateNirSection7(),
                CreateNirSection8(),
                CreateNirSection9()
            };
        }

        private static NirSectionSeed CreateNirTitleSection() => new NirSectionSeed
        {
            Number = 0,
            Title = "Титульный лист"
        };

        private static NirSectionSeed CreateNirSection1() => new NirSectionSeed
        {
            Number = 1,
            Title = "Договорная деятельность",
            SubSections = new[]
            {
                new NirSubSectionSeed
                {
                    Number = 1,
                    Title = "Общая характеристика научных исследований",
                    TextPatternName = "nir_1_1_text",
                    TextContent = "Дать количество договоров и объем финансирования. Отразить отдельно количество и выполнение НИР по Международным, Федеральным, Региональным программам и проектам, по распоряжению ОАО «РЖД», по постановлению РАН (с указанием участия в качестве головного; ответственного по направлению или исполнителя), по грантам на научные исследования (международного, федерального уровня) с указанием объема и источника финансирования."
                },
                new NirSubSectionSeed
                {
                    Number = 2,
                    Title = "Хоздоговорные НИР",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Хоздоговорные научно-исследовательские работы кафедры (НИЦ, НИЛ, НИГ)",
                            PatternName = "nir_1_2_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ договора"),
                                HeaderCell(1, 2, "Наименование (основание к ее выполнению: распоряжение ОАО «РЖД», научно-техническая программа и т.п.) и Ф.И.О. руководителя"),
                                HeaderCell(1, 3, "Стоимость НИР в 2о… г."),
                                HeaderCell(1, 4, "Заказчик"),
                                HeaderCell(1, 5, "Фактическое выполнение НИР, тыс. руб.", 2),
                                HeaderCell(1, 7, "Краткая характеристика выполненных разделов темы"),
                                HeaderCell(2, 5, "Акты"),
                                HeaderCell(2, 6, "Оплата")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 3,
                    Title = "Темы госбюджетных научных исследований",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Темы госбюджетных научных исследований",
                            PatternName = "nir_1_3_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ п/п"),
                                HeaderCell(1, 2, "Номер в Реестре тем ГБНИ"),
                                HeaderCell(1, 3, "Наименование НИР,\nФ.И.О. руководителя"),
                                HeaderCell(1, 4, "Сроки выполнения", 2),
                                HeaderCell(1, 6, "Краткая характеристика выполненной работы в 2024 г."),
                                HeaderCell(1, 7, "Кому представлен отчет. Место и характеристика использования НИР"),
                                HeaderCell(2, 4, "начало"),
                                HeaderCell(2, 5, "окончание")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 4,
                    Title = "Реализация результатов НИОКР",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Реализация результатов НИОКР",
                            PatternName = "nir_1_4_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ п/п"),
                                HeaderCell(1, 2, "Шифр темы"),
                                HeaderCell(1, 3, "Наименование темы"),
                                HeaderCell(1, 4, "Заказчик.\nМесто внедрения"),
                                HeaderCell(1, 5, "Ожидаемая технико-экономическая или иная эффективность"),
                                HeaderCell(1, 6, "Форма внедрения (опытный образец, серийное производство, отрасл. инстр., нормативная документация и т.п.)")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 5,
                    Title = "Международные творческие связи и содружество с производством",
                    TextPatternName = "nir_1_5_text",
                    TextContent = "На кафедре имеется базовая кафедра в Новосибирском информационно-вычислительном центре, специалисты которого ведут дисциплины «Информационные системы железнодорожного транспорта» для направления 09.03.02 Информационные системы и технологии, а также «Корпоративные информационные системы на железнодорожном транспорте» для направления 09.03.03 Прикладная информатика. Начальник Новосибирского ИВЦ Шабанов А.Н., главный инженер Шелеметев И.В., а также менеджер по информационным технологиям службы развития процессного управления и управления качеством ГВЦ - филиала ОАО «РЖД» Пинигин С.Ю. являются членами ГЭК по двум направлениям бакалавриата, а также в магистратуре по направлению 09.04.02. Кроме того, начальник аспирантуры ФГУП НИИ СибНИА Железнов Л.П. ведет дисциплину Проектирование информационных систем по направлению 09.03.02 и дисциплину «Методы проектирования информационных систем» по направлению 09.03.03. Ведущий инженер-проектировщик центра финансовых технологий Силкачева О.Н. ведет дисциплину «Информационная безопасность» по направлениям 09.03.02 и 09.03.03. Специалисты ООО «Информационные системы и сервисы» Вавилова С.А., Ксенофонтова М.А. ведут дисциплину «Моделирование бизнес-процессов»."
                }
            }
        };

        private static NirSectionSeed CreateNirSection2() => new NirSectionSeed
        {
            Number = 2,
            Title = "Остепененность и подготовка научных кадров",
            SubSections = new[]
            {
                new NirSubSectionSeed
                {
                    Number = 1,
                    Title = "АСПИРАНТУРА",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "АСПИРАНТУРА",
                            "nir_2_1_table",
                            "№ п/п",
                            "Ф.И.О. аспиранта",
                            "Ф.И.О. научного руководителя",
                            "Наименование темы диссертации",
                            "Участие в хоздоговорной работе (+ / -);\nуказать подразделение",
                            "Участие в учебном процессе (+/-);\nуказать подразделение")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 2,
                    Title = "ДОКТОРАНТУРА",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "ДОКТОРАНТУРА",
                            "nir_2_2_table",
                            "№ п/п",
                            "Ф.И.О. докторанта",
                            "Ф.И.О. научного консультанта",
                            "Наименование темы диссертационной работы",
                            "Шифр и наименование научной специальности",
                            "Наименование принимающей организации")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 3,
                    Title = "ЗАЩИТЫ",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "ЗАЩИТЫ",
                            "nir_2_3_table",
                            "№ п/п",
                            "Ф.И.О. сотрудника (аспиранта), защитившего диссертацию",
                            "Ф.И.О. научного руководителя (консультанта)",
                            "Дата защиты в диссертационном совете",
                            "Дата и номер приказа Минобрнауки РФ о выдаче диплома")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 4,
                    Title = "ПЛАН ЗАЩИТ",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "ПЛАН ЗАЩИТ",
                            "nir_2_4_table",
                            "№ п/п",
                            "Ф.И.О. сотрудника (аспиранта)",
                            "Ф.И.О. научного руководителя (консультанта)",
                            "Планируемая дата подачи документов в диссертационный совет",
                            "Наименование организации, при которой создан диссертационный совет")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 5,
                    Title = "ДИССЕРТАЦИОННЫЕ СОВЕТЫ",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "ДИССЕРТАЦИОННЫЕ СОВЕТЫ",
                            "nir_2_5_table",
                            "№ п/п",
                            "Ф.И.О. члена совета",
                            "Шифр диссертационного совета",
                            "Шифр и наименование научной специальности",
                            "Наименование организации, при которой создан диссертационный совет")
                    }
                }
            }
        };
        private static NirSectionSeed CreateNirSection3() => new NirSectionSeed
        {
            Number = 3,
            Title = "Издательская деятельность",
            SubSections = new[]
            {
                new NirSubSectionSeed
                {
                    Number = 1,
                    Title = "Издательская деятельность сотрудников кафедры и лаборатории",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "Издательская деятельность сотрудников кафедры и лаборатории",
                            "nir_3_1_table",
                            "№",
                            "Доля авторов заполняющей кафедры",
                            "Фамилия И.О. авторов",
                            "Наименование публикации",
                            "Тип публикации",
                            "Наименование издания, год, страницы публикации (без кавычек)",
                            "Место издания (наименование организации, город)")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 2,
                    Title = "Публикационная активность профессорско-преподавательского состава кафедры и научных сотрудников лаборатории",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Публикационная активность профессорско-преподавательского состава кафедры и научных сотрудников лаборатории",
                            PatternName = "nir_3_2_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ п/п", 1, 2),
                                HeaderCell(1, 2, "Ф.И.О. автора", 1, 2),
                                HeaderCell(1, 3, "Количество публикаций в отчетном году", 3),
                                HeaderCell(1, 6, "Количество цитирований в отчетном году", 4),
                                HeaderCell(2, 3, "RSCI+\nScopus+Wos⁴"),
                                HeaderCell(2, 4, "РИНЦ\nжурналы"),
                                HeaderCell(2, 5, "РИНЦ\nдругие\nиздания"),
                                HeaderCell(2, 6, "WOS"),
                                HeaderCell(2, 7, "Scopus"),
                                HeaderCell(2, 8, "РИНЦ\nс учетом\nсамоцит."),
                                HeaderCell(2, 9, "РИНЦ\nбез учета\nсамоцит.")
                            }
                        }
                    }
                }
            }
        };

        private static NirSectionSeed CreateNirSection4() => new NirSectionSeed
        {
            Number = 4,
            Title = "Регистрация объектов интеллектуальной собственности",
            TextPatternName = "nir_4_text",
            TextContent = "Привести сравнительный анализ плановых показателей, заявленных на 2024 г. с фактическим выполнением.",
            SubSections = new[]
            {
                new NirSubSectionSeed
                {
                    Number = 1,
                    Title = "Общие сведения",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Общие сведения",
                            PatternName = "nir_4_1_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 2, "План"),
                                HeaderCell(1, 3, "Факт"),
                                BodyCell(2, 1, "Количество заявок на изобретения"),
                                BodyCell(3, 1, "Количество заявок на полезную модель"),
                                BodyCell(4, 1, "Количество заявок на программные продукты и базы данных, электронные ресурсы")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 2,
                    Title = "Полученные патенты/положительные решения на изобретения и полезные модели",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "Полученные патенты/положительные решения на изобретения и полезные модели",
                            "nir_4_2_table",
                            "Автор (-ы)",
                            "№ патента",
                            "Наименование заявки",
                            "Правообладатель")
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 3,
                    Title = "Полученные свидетельства на программные продукты и базы данных, электронные ресурсы",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "Полученные свидетельства на программные продукты и базы данных, электронные ресурсы",
                            "nir_4_3_table",
                            "Автор (-ы)",
                            "№ свидетельства",
                            "Наименование программного продукта, электронного ресурса",
                            "Правообладатель")
                    }
                }
            }
        };

        private static NirSectionSeed CreateNirSection5() => new NirSectionSeed
        {
            Number = 5,
            Title = "Инновационная деятельность",
            SubSections = new[]
            {
                new NirSubSectionSeed
                {
                    Number = 1,
                    Title = "Реализуемые стартап-проекты",
                    Tables = new[]
                    {
                        CreateSimpleTableSeed(
                            "Реализуемые стартап-проекты",
                            "nir_5_1_table",
                            "№ п/п",
                            "Тема проекта",
                            "Команда проекта (руководитель, исполнители)",
                            "Период реализации\n(мес.год.- мес.год.)",
                            "Потенциальный заказчик",
                            "Источник финансирования",
                            "Наличие дорожной карты\n(+/-)")
                    },
                    Children = new[]
                    {
                        new NirSubSectionSeed
                        {
                            Number = 1,
                            Title = "Планируемые стартап-проекты",
                            Tables = new[]
                            {
                                CreateSimpleTableSeed(
                                    "Планируемые стартап-проекты",
                                    "nir_5_1_1_table",
                                    "№ п/п",
                                    "Тема проекта",
                                    "Команда проекта (руководитель, исполнители)",
                                    "Потенциальный заказчик",
                                    "Источник финансирования")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 2,
                    Title = "Участие в конкурсах, грантах",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Участие в конкурсах, грантах",
                            PatternName = "nir_5_2_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ п/п"),
                                HeaderCell(1, 2, "Наименование конкурса"),
                                HeaderCell(1, 3, "Наименование заявки (темы) и состав исполнителей"),
                                HeaderCell(1, 4, "Общий объем финансирования / объем освоенных средств за отчетный период, тыс. руб."),
                                HeaderCell(1, 5, "Краткая характеристика выполненной работы"),
                                BodyCell(2, 1, "Конкурсы на соискание грантов", 5),
                                BodyCell(3, 1, "Иные конкурсы", 5)
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 3,
                    Title = "Научные конференции",
                    TextPatternName = "nir_5_3_text",
                    TextContent = "Участие и организация научно-технических конференций, семинаров (необходимо дать название и дату). Кафедра участвовала в конференции «Политранспортные системы». Работала 1 секция, заслушано 18 докладов. Статьи отданы в издательство.",
                    Children = new[]
                    {
                        new NirSubSectionSeed
                        {
                            Number = 1,
                            Title = "Организованные конференции",
                            Tables = new[]
                            {
                                CreateSimpleTableSeed(
                                    "Организованные конференции",
                                    "nir_5_3_1_table",
                                    "№ п/п",
                                    "Наименование конференции, семинара\n(привести название и дату)",
                                    "Работало секций",
                                    "Общее число участников конференции",
                                    "Заслушано докладов",
                                    "Сборник по итогам мероприятия\n(название, кол-во печ. л.)")
                            }
                        },
                        new NirSubSectionSeed
                        {
                            Number = 3,
                            Title = "План проведения конференций, семинаров, совещаний в 2025 г.",
                            Tables = new[]
                            {
                                CreateSimpleTableSeed(
                                    "План проведения конференций, семинаров, совещаний в 2025 г.",
                                    "nir_5_3_3_table",
                                    "№ п/п",
                                    "Наименование мероприятия",
                                    "Дата проведения",
                                    "Ответственный от кафедры",
                                    "Формат публикации результатов")
                            }
                        }
                    }
                },
                new NirSubSectionSeed
                {
                    Number = 4,
                    Title = "Участие в выставках-ярмарках",
                    Tables = new[]
                    {
                        new TableSeed
                        {
                            Title = "Участие в выставках-ярмарках",
                            PatternName = "nir_5_4_table",
                            Cells = new[]
                            {
                                HeaderCell(1, 1, "№ п/п"),
                                HeaderCell(1, 2, "Наименование\nвыставки"),
                                HeaderCell(1, 3, "Наименование\nэкспоната"),
                                HeaderCell(1, 4, "Получено наград за 1, 2, 3 место\n(медаль, диплом, грамота)"),
                                BodyCell(2, 1, "Научные проекты и разработки", 4),
                                BodyCell(3, 1, "Книжная продукция", 4)
                            }
                        }
                    }
                }
            }
        };
        private static NirSectionSeed CreateNirSection6() => new NirSectionSeed
        {
            Number = 6,
            Title = "Развитие учебно-лабораторной и научной базы",
            Tables = new[]
            {
                new TableSeed
                {
                    Title = "Развитие учебно-лабораторной и научной базы",
                    PatternName = "nir_6_table",
                    Cells = new[]
                    {
                        HeaderCell(1, 1, "Перечень приобретенной продукции", 1, 3),
                        HeaderCell(1, 2, "Финансирование, тыс. руб.", 3),
                        HeaderCell(2, 2, "Для учебных целей", 2),
                        HeaderCell(2, 4, "Для научных подразделений", 1, 2),
                        HeaderCell(3, 2, "за счет г/б"),
                        HeaderCell(3, 3, "за счет научных средств"),
                        BodyCell(4, 1, "Учебно-лабораторное оборудование"),
                        BodyCell(5, 1, "Оргтехника"),
                        BodyCell(6, 1, "Комплектующие"),
                        BodyCell(7, 1, "Программный продукт"),
                        BodyCell(8, 1, "Технологии"),
                        BodyCell(9, 1, "И т.д."),
                        BodyCell(10, 1, "Всего")
                    }
                }
            }
        };

        private static NirSectionSeed CreateNirSection7() => new NirSectionSeed
        {
            Number = 7,
            Title = "Выводы и рекомендации",
            TextPatternName = "nir_7_text",
            TextContent = "Здесь будут описаны выводы и рекомендации."
        };

        private static NirSectionSeed CreateNirSection8() => new NirSectionSeed
        {
            Number = 8,
            Title = "Обобщенные показатели научной деятельности кафедры",
            Tables = new[]
            {
                new TableSeed
                {
                    Title = "Обобщенные показатели научной деятельности кафедры",
                    PatternName = "nir_8_table",
                    Cells = new[]
                    {
                        HeaderCell(1, 1, "№ п/п"),
                        HeaderCell(1, 2, "Наименование показателя"),
                        HeaderCell(1, 3, "2023 г."),
                        HeaderCell(1, 4, "2024 г."),
                        BodyCell(2, 1, "1"),
                        BodyCell(2, 2, "«Остепененность» по кафедре:"),
                        BodyCell(3, 2, "расчетная остепененность %"),
                        BodyCell(4, 2, "докторов наук, профессоров"),
                        BodyCell(5, 1, "2"),
                        BodyCell(5, 2, "Количество аспирантов по кафедре"),
                        BodyCell(6, 1, "3"),
                        BodyCell(6, 2, "Количество докторантов по кафедре"),
                        BodyCell(7, 1, "4"),
                        BodyCell(7, 2, "Защита диссертаций"),
                        BodyCell(8, 2, "диссертации, защищенных под научным руководством сотрудников кафедры"),
                        BodyCell(9, 2, "докторских сотрудниками кафедры"),
                        BodyCell(10, 2, "кандидатских сотрудниками кафедры"),
                        BodyCell(11, 1, "5"),
                        BodyCell(11, 2, "Договорная деятельность"),
                        BodyCell(12, 2, "Объем выполненных работ, тыс. руб., в том числе:"),
                        BodyCell(13, 2, "- объем хоздоговорных работ, тыс. руб."),
                        BodyCell(14, 2, "- объем внешних грантов, тыс. руб."),
                        BodyCell(15, 2, "Госбюджетное научное исследование"),
                        BodyCell(16, 2, "Стартап-проект (диплом, как стартап)"),
                        BodyCell(17, 1, "6"),
                        BodyCell(17, 2, "Патентная деятельность"),
                        BodyCell(18, 2, "Получено патентов"),
                        BodyCell(19, 2, "Получено свидетельств на ПО"),
                        BodyCell(20, 2, "Подано заявок на патенты"),
                        BodyCell(21, 1, "7"),
                        BodyCell(21, 2, "Издательская деятельность"),
                        BodyCell(22, 2, "Учебники, учебные пособия"),
                        BodyCell(23, 2, "Монографии"),
                        BodyCell(24, 2, "Всего публикаций (без учета тезисов), в том числе:"),
                        BodyCell(25, 2, "- в изданиях, входящих в журналы RSCI+ Scopus +WoS"),
                        BodyCell(26, 2, "- в журналах из перечня ВАК"),
                        BodyCell(27, 2, "- в журналах, входящем в РИНЦ (исключая журналы из перечня ВАК)"),
                        BodyCell(28, 2, "- в изданиях, входящих в РИНЦ (исключая учтенные в иных строках)"),
                        BodyCell(29, 2, "- публикации в других изданиях"),
                        BodyCell(30, 2, "Суммарное цитирование работ сотрудников кафедры в РИНЦ (без учета самоцитирования)"),
                        BodyCell(31, 2, "Суммарное цитирование статей, индексируемых в международных базах данных Web of Science / Scopus"),
                        BodyCell(32, 1, "8"),
                        BodyCell(32, 2, "Организация и выступления на конференциях"),
                        BodyCell(33, 2, "Организация и проведение конференций"),
                        BodyCell(34, 2, "Организация и проведение семинаров"),
                        BodyCell(35, 2, "Выступление на международной конференции"),
                        BodyCell(36, 2, "Выступление на всероссийской (национальной, всероссийской с международным участием) конференции"),
                        BodyCell(37, 2, "Выступление на региональной и областной конференции"),
                        BodyCell(38, 2, "Выступление на локальной (городской, межвузовской, внутривузовской) конференции"),
                        BodyCell(39, 1, "9"),
                        BodyCell(39, 2, "Выставки"),
                        BodyCell(40, 2, "Количество выставок"),
                        BodyCell(41, 2, "Представлено экспонатов"),
                        BodyCell(42, 2, "Получено наград за 1, 2, 3 место (медаль, диплом, грамота)"),
                        BodyCell(43, 1, "10"),
                        BodyCell(43, 2, "Конкурсы"),
                        BodyCell(44, 2, "Количество конкурсов"),
                        BodyCell(45, 2, "Представлено работ"),
                        BodyCell(46, 2, "Получено наград"),
                        BodyCell(47, 1, "11"),
                        BodyCell(47, 2, "Участие в работе диссертационных советов"),
                        BodyCell(48, 2, "Председатель"),
                        BodyCell(49, 2, "Ученый секретарь диссертационного совета"),
                        BodyCell(50, 2, "Член совета"),
                        BodyCell(51, 1, "12"),
                        BodyCell(51, 2, "Условный штат кафедры и НИЛ (НИГ, НИЦ)"),
                        BodyCell(52, 2, "Условный штат кафедры"),
                        BodyCell(53, 2, "Ставки НИЛ (НИГ, НИЦ)")
                    }
                }
            }
        };
        private static NirSectionSeed CreateNirSection9() => new NirSectionSeed
        {
            Number = 9,
            Title = "Раздел 9"
        };
    }
}
