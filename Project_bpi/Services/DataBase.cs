using Project_bpi.Models;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace Project_bpi.Services
{
    public class DataBase
    {
        private readonly string _connectionString;
        private readonly string databasePath;
        
        public string DatabasePath => databasePath;
        
        public DataBase(string databaseFileName = "Kurs.db")
        {
            databasePath = Path.IsPathRooted(databaseFileName)
                ? databaseFileName
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, databaseFileName);
            _connectionString = $"Data Source={databasePath};";
        }
        // Запуск Базы данных
        public void InitializeDatabase(bool showMessageWhenCreated = true)
        {
            bool databaseExists = File.Exists(databasePath);
            if (!databaseExists)
            {
                var directory = Path.GetDirectoryName(databasePath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.Create(databasePath).Close();
                if (showMessageWhenCreated)
                {
                    MessageBox.Show("Файл базы данных не был обнаружен, создан новый файл базы данных.");
                }
            }
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                using (var command = new SQLiteCommand("PRAGMA foreign_keys = ON;", connection))
                {
                    command.ExecuteNonQuery();
                }
                string createTablesScript = @"
                CREATE TABLE IF NOT EXISTS ""FilePatterns"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Title"" TEXT NOT NULL,
                    ""Year"" INTEGER NOT NULL,
                    ""Path"" TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS ""Report"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Title"" TEXT NOT NULL,
                    ""Year"" INTEGER NOT NULL,
                    ""Pattern_id"" INTEGER NOT NULL,
                    FOREIGN KEY(""Pattern_id"") REFERENCES ""FilePatterns""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Section"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Report_id"" INTEGER NOT NULL,
                    ""Number"" INTEGER NOT NULL,
                    ""Title"" TEXT,
                    FOREIGN KEY(""Report_id"") REFERENCES ""Report""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Subsection"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Section_id"" INTEGER NOT NULL,
                    ""ParentSubsection_id"" INTEGER NULL,
                    ""Number"" INTEGER NOT NULL,
                    ""Title"" TEXT,
                    FOREIGN KEY(""Section_id"") REFERENCES ""Section""(""id"") ON DELETE CASCADE,
                    FOREIGN KEY(""ParentSubsection_id"") REFERENCES ""Subsection""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Table"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Title"" TEXT NOT NULL,
                    ""Subsection_id"" INTEGER NOT NULL,
                    ""Pattern_name"" TEXT NOT NULL,
                    FOREIGN KEY(""Subsection_id"") REFERENCES ""Subsection""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Table_item"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Table_id"" INTEGER NOT NULL,
                    ""Column"" INTEGER NOT NULL,
                    ""Row"" INTEGER NOT NULL,
                    ""Header"" TEXT NOT NULL,
                    ""ColSpan"" INTEGER NOT NULL DEFAULT 1,
                    ""RowSpan"" INTEGER NOT NULL DEFAULT 1,
                    ""IsHeader"" INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY(""Table_id"") REFERENCES ""Table""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Text"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Subsection_id"" INTEGER NOT NULL,
                    ""Content"" TEXT NOT NULL,
                    ""Pattern_name"" TEXT NOT NULL,
                    FOREIGN KEY(""Subsection_id"") REFERENCES ""Subsection""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""ChangeHistory"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""ChangedAtUtc"" TEXT NOT NULL,
                    ""ActionType"" TEXT NOT NULL,
                    ""EntityType"" TEXT NOT NULL,
                    ""Location"" TEXT NOT NULL,
                    ""Details"" TEXT NULL
                );

                CREATE TABLE IF NOT EXISTS ""SuppressedBuiltInReports"" (
                    ""Key"" TEXT NOT NULL PRIMARY KEY
                );";
                using (var command = new SQLiteCommand(createTablesScript, connection))
                {
                    command.ExecuteNonQuery();
                }

                EnsureColumnExists(connection, "Table_item", "ColSpan", @"ALTER TABLE ""Table_item"" ADD COLUMN ""ColSpan"" INTEGER NOT NULL DEFAULT 1");
                EnsureColumnExists(connection, "Table_item", "RowSpan", @"ALTER TABLE ""Table_item"" ADD COLUMN ""RowSpan"" INTEGER NOT NULL DEFAULT 1");
                EnsureColumnExists(connection, "Table_item", "IsHeader", @"ALTER TABLE ""Table_item"" ADD COLUMN ""IsHeader"" INTEGER NOT NULL DEFAULT 0");
                EnsureColumnExists(connection, "Subsection", "ParentSubsection_id", @"ALTER TABLE ""Subsection"" ADD COLUMN ""ParentSubsection_id"" INTEGER NULL");
            }
        }

        private void EnsureColumnExists(SQLiteConnection connection, string tableName, string columnName, string alterSql)
        {
            using (var command = new SQLiteCommand($@"PRAGMA table_info(""{tableName}"");", connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (string.Equals(reader["name"]?.ToString(), columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
            }

            using (var alterCommand = new SQLiteCommand(alterSql, connection))
            {
                alterCommand.ExecuteNonQuery();
            }
        }

        // Добавление в бд нового ПУТИ файла-шаблона
        public async Task<int> AddFilePattern(PatternFile pattern)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO FilePatterns (Title, Year, Path) 
                            VALUES (@Title, @Year, @Path); 
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Title", pattern.Title);
                    command.Parameters.AddWithValue("@Year", pattern.Year);
                    command.Parameters.AddWithValue("@Path", pattern.Path);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task<int> AddReport(Report report)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO Report (Title, Year, Pattern_id)
                            VALUES (@Title, @Year, @PatternId);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Title", report.Title);
                    command.Parameters.AddWithValue("@Year", report.Year);
                    command.Parameters.AddWithValue("@PatternId", report.PattarnId);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateFilePattern(PatternFile pattern)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE FilePatterns
                            SET Title = @Title, Year = @Year, Path = @Path
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", pattern.Id);
                    command.Parameters.AddWithValue("@Title", pattern.Title);
                    command.Parameters.AddWithValue("@Year", pattern.Year);
                    command.Parameters.AddWithValue("@Path", pattern.Path);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task UpdateReport(Report report)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE Report
                            SET Title = @Title, Year = @Year, Pattern_id = @PatternId
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", report.Id);
                    command.Parameters.AddWithValue("@Title", report.Title);
                    command.Parameters.AddWithValue("@Year", report.Year);
                    command.Parameters.AddWithValue("@PatternId", report.PattarnId);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task<int> AddHistoryEntry(HistoryEntry entry)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                const string query = @"INSERT INTO ""ChangeHistory"" (ChangedAtUtc, ActionType, EntityType, Location, Details)
                                       VALUES (@ChangedAtUtc, @ActionType, @EntityType, @Location, @Details);
                                       SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ChangedAtUtc", entry.ChangedAtUtc.ToUniversalTime().ToString("o"));
                    command.Parameters.AddWithValue("@ActionType", entry.ActionType ?? string.Empty);
                    command.Parameters.AddWithValue("@EntityType", entry.EntityType ?? string.Empty);
                    command.Parameters.AddWithValue("@Location", entry.Location ?? string.Empty);
                    command.Parameters.AddWithValue("@Details", string.IsNullOrWhiteSpace(entry.Details)
                        ? (object)DBNull.Value
                        : entry.Details);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task<List<HistoryEntry>> GetHistoryEntries(int limit = 500)
        {
            var entries = new List<HistoryEntry>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                const string query = @"SELECT id, ChangedAtUtc, ActionType, EntityType, Location, Details
                                       FROM ""ChangeHistory""
                                       ORDER BY datetime(ChangedAtUtc) DESC, id DESC
                                       LIMIT @Limit";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Limit", limit);

                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            string changedAtRaw = reader.GetString(1);
                            DateTime changedAtUtc = DateTime.TryParse(
                                changedAtRaw,
                                null,
                                System.Globalization.DateTimeStyles.RoundtripKind,
                                out var parsedChangedAt)
                                ? parsedChangedAt.ToUniversalTime()
                                : DateTime.UtcNow;

                            entries.Add(new HistoryEntry
                            {
                                Id = reader.GetInt32(0),
                                ChangedAtUtc = changedAtUtc,
                                ActionType = reader.GetString(2),
                                EntityType = reader.GetString(3),
                                Location = reader.GetString(4),
                                Details = reader.IsDBNull(5) ? string.Empty : reader.GetString(5)
                            });
                        }
                    }
                }
            }

            return entries;
        }

        public async Task<bool> HasSuppressedBuiltInReport(string key)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                const string query = @"SELECT COUNT(1) FROM ""SuppressedBuiltInReports"" WHERE ""Key"" = @Key";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Key", key ?? string.Empty);
                    return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
                }
            }
        }

        public async Task AddSuppressedBuiltInReport(string key)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                const string query = @"INSERT OR IGNORE INTO ""SuppressedBuiltInReports"" (""Key"") VALUES (@Key)";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Key", key ?? string.Empty);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        // Удаление шаблона
        public async Task DeleteFilePattern(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"DELETE from FilePatterns where id = @id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@id", id);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task<List<Report>> GetAllReports()
        {
            var reports = new List<Report>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "SELECT id, Title, Year, Pattern_id FROM Report ORDER BY Year DESC, Title";
                using (var command = new SQLiteCommand(query, connection))
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        reports.Add(new Report
                        {
                            Id = reader.GetInt32(0),
                            Title = reader.GetString(1),
                            Year = reader.GetInt32(2),
                            PattarnId = reader.GetInt32(3),
                            Sections = new List<Section>()
                        });
                    }
                }
            }

            return reports;
        }

        // Вывод уже готового отчётаы
        public async Task<Report> GetFullReport(int reportId)
        {
            Report report = null;

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string reportQuery = @"SELECT r.id, r.Title, r.Year, r.Pattern_id,
                                   fp.Title, fp.Year, fp.Path
                                   FROM Report r
                                   INNER JOIN FilePatterns fp ON r.Pattern_id = fp.id
                                   WHERE r.id = @ReportId";

                using (var command = new SQLiteCommand(reportQuery, connection))
                {
                    command.Parameters.AddWithValue("@ReportId", reportId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        if (await reader.ReadAsync())
                        {
                            report = new Report
                            {
                                Id = reader.GetInt32(0),
                                Title = reader.GetString(1),
                                Year = reader.GetInt32(2),
                                PattarnId = reader.GetInt32(3),
                                PatternFile = new PatternFile
                                {
                                    Id = reader.GetInt32(3),
                                    Title = reader.GetString(4),
                                    Year = reader.GetInt32(5),
                                    Path = reader.GetString(6)
                                },
                                Sections = new List<Section>()
                            };
                        }
                    }
                }

                if (report != null)
                {
                    await LoadSectionsAsync(connection, report);
                }
            }

            return report;
        }

        private async Task LoadSectionsAsync(SQLiteConnection connection, Report report)
        {
            string sectionsQuery = @"SELECT id, Report_id, Number, Title
                                     FROM Section
                                     WHERE Report_id = @ReportId
                                     ORDER BY Number";

            using (var command = new SQLiteCommand(sectionsQuery, connection))
            {
                command.Parameters.AddWithValue("@ReportId", report.Id);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        report.Sections.Add(new Section
                        {
                            Id = reader.GetInt32(0),
                            ReportId = reader.GetInt32(1),
                            Number = reader.GetInt32(2),
                            Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                            Report = report,
                            SubSections = new List<SubSection>()
                        });
                    }
                }
            }

            foreach (var section in report.Sections)
            {
                await LoadSubSectionsAsync(connection, section);
            }
        }

        private async Task LoadSubSectionsAsync(SQLiteConnection connection, Section section)
        {
            var allSubSections = new List<SubSection>();
            var subSectionMap = new Dictionary<int, SubSection>();

            string subsectionsQuery = @"SELECT id, Section_id, ParentSubsection_id, Number, Title
                                        FROM Subsection
                                        WHERE Section_id = @SectionId
                                        ORDER BY COALESCE(ParentSubsection_id, 0), Number, id";

            using (var command = new SQLiteCommand(subsectionsQuery, connection))
            {
                command.Parameters.AddWithValue("@SectionId", section.Id);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        var subsection = new SubSection
                        {
                            Id = reader.GetInt32(0),
                            SectionId = reader.GetInt32(1),
                            ParentSubsectionId = reader.IsDBNull(2) ? (int?)null : reader.GetInt32(2),
                            Number = reader.GetInt32(3),
                            Title = reader.IsDBNull(4) ? null : reader.GetString(4),
                            Section = section,
                            SubSections = new List<SubSection>(),
                            Tables = new List<Table>(),
                            Texts = new List<Text>()
                        };

                        allSubSections.Add(subsection);
                        subSectionMap[subsection.Id] = subsection;
                    }
                }
            }

            foreach (var subsection in allSubSections)
            {
                if (subsection.ParentSubsectionId.HasValue &&
                    subSectionMap.TryGetValue(subsection.ParentSubsectionId.Value, out var parentSubsection))
                {
                    subsection.ParentSubsection = parentSubsection;
                    parentSubsection.SubSections.Add(subsection);
                }
                else
                {
                    section.SubSections.Add(subsection);
                }
            }

            foreach (var subsection in allSubSections)
            {
                await LoadSubSectionContentsAsync(connection, subsection);
            }
        }

        private async Task LoadSubSectionContentsAsync(SQLiteConnection connection, SubSection subsection)
        {
            string tablesQuery = @"SELECT id, Title, Subsection_id, Pattern_name
                                   FROM ""Table""
                                   WHERE Subsection_id = @SubsectionId";

            using (var command = new SQLiteCommand(tablesQuery, connection))
            {
                command.Parameters.AddWithValue("@SubsectionId", subsection.Id);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        subsection.Tables.Add(new Table
                        {
                            Id = reader.GetInt32(0),
                            Title = reader.GetString(1),
                            SubsectionId = reader.GetInt32(2),
                            PatternName = reader.GetString(3),
                            TableItems = new List<TableItem>()
                        });
                    }
                }
            }

            string textsQuery = @"SELECT id, Subsection_id, Content, Pattern_name
                                  FROM ""Text""
                                  WHERE Subsection_id = @SubsectionId";

            using (var command = new SQLiteCommand(textsQuery, connection))
            {
                command.Parameters.AddWithValue("@SubsectionId", subsection.Id);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        subsection.Texts.Add(new Text
                        {
                            Id = reader.GetInt32(0),
                            SubsectionId = reader.GetInt32(1),
                            Content = reader.GetString(2),
                            PatternName = reader.GetString(3)
                        });
                    }
                }
            }

            foreach (var table in subsection.Tables)
            {
                await LoadTableItemsAsync(connection, table);
            }
        }

        private async Task LoadTableItemsAsync(SQLiteConnection connection, Table table)
        {
            string itemsQuery = @"SELECT id, Table_id, Column, Row, Header, ColSpan, RowSpan, IsHeader
                                  FROM ""Table_item""
                                  WHERE Table_id = @TableId
                                  ORDER BY IsHeader DESC, Row, Column";

            using (var command = new SQLiteCommand(itemsQuery, connection))
            {
                command.Parameters.AddWithValue("@TableId", table.Id);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        table.TableItems.Add(new TableItem
                        {
                            Id = reader.GetInt32(0),
                            TableId = reader.GetInt32(1),
                            Column = reader.GetInt32(2),
                            Row = reader.GetInt32(3),
                            Header = reader.GetString(4),
                            ColSpan = reader.IsDBNull(5) ? 1 : reader.GetInt32(5),
                            RowSpan = reader.IsDBNull(6) ? 1 : reader.GetInt32(6),
                            IsHeader = !reader.IsDBNull(7) && reader.GetInt32(7) == 1
                        });
                    }
                }
            }
        }

        // ----------------------------------------------------------------------
        // Работа с таблицей
        public async Task<int> AddTable(Table table)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO ""Table"" (Title, Subsection_id, Pattern_name) 
                            VALUES (@Title, @SubsectionId, @PatternName);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Title", table.Title);
                    command.Parameters.AddWithValue("@SubsectionId", table.SubsectionId);
                    command.Parameters.AddWithValue("@PatternName", table.PatternName);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateTable(Table table)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE ""Table"" 
                            SET Title = @Title, Pattern_name = @PatternName 
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", table.Id);
                    command.Parameters.AddWithValue("@Title", table.Title);
                    command.Parameters.AddWithValue("@PatternName", table.PatternName);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteTable(int tableId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM \"Table\" WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", tableId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        // ----------------------------------------------------------------------
        // Работа с ячейками таблицы
        public async Task<int> AddTableItem(TableItem item)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO ""Table_item"" (Table_id, Column, Row, Header, ColSpan, RowSpan, IsHeader) 
                            VALUES (@TableId, @Column, @Row, @Header, @ColSpan, @RowSpan, @IsHeader);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableId", item.TableId);
                    command.Parameters.AddWithValue("@Column", item.Column);
                    command.Parameters.AddWithValue("@Row", item.Row);
                    command.Parameters.AddWithValue("@Header", item.Header);
                    command.Parameters.AddWithValue("@ColSpan", item.ColSpan);
                    command.Parameters.AddWithValue("@RowSpan", item.RowSpan);
                    command.Parameters.AddWithValue("@IsHeader", item.IsHeader ? 1 : 0);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task DeleteAllTableItems(int tableId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM \"Table_item\" WHERE Table_id = @TableId";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableId", tableId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }
        // --------------------------------------------------------------------------
        // Работа с текстовыми данными
        public async Task<int> AddText(Text text)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO ""Text"" (Subsection_id, Content, Pattern_name) 
                            VALUES (@SubsectionId, @Content, @PatternName);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@SubsectionId", text.SubsectionId);
                    command.Parameters.AddWithValue("@Content", text.Content);
                    command.Parameters.AddWithValue("@PatternName", text.PatternName);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateText(Text text)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE ""Text"" 
                            SET Content = @Content, Pattern_name = @PatternName 
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", text.Id);
                    command.Parameters.AddWithValue("@Content", text.Content);
                    command.Parameters.AddWithValue("@PatternName", text.PatternName);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteText(int textId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM \"Text\" WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", textId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        // -----------------------------------------------------------------------------
        // работа с подразрелами
        public async Task<int> AddSubsection(SubSection subsection)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO Subsection (Section_id, ParentSubsection_id, Number, Title) 
                            VALUES (@SectionId, @ParentSubsectionId, @Number, @Title);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@SectionId", subsection.SectionId);
                    command.Parameters.AddWithValue("@ParentSubsectionId", subsection.ParentSubsectionId ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Number", subsection.Number);
                    command.Parameters.AddWithValue("@Title", subsection.Title ?? (object)DBNull.Value);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateSubsection(SubSection subsection)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE Subsection 
                            SET ParentSubsection_id = @ParentSubsectionId, Number = @Number, Title = @Title 
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", subsection.Id);
                    command.Parameters.AddWithValue("@ParentSubsectionId", subsection.ParentSubsectionId ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Number", subsection.Number);
                    command.Parameters.AddWithValue("@Title", subsection.Title ?? (object)DBNull.Value);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteSubsection(int subsectionId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                using (var transaction = connection.BeginTransaction())
                {
                    await DeleteSubsectionInternalAsync(connection, transaction, subsectionId);
                    transaction.Commit();
                }
            }
        }

        private async Task DeleteSubsectionInternalAsync(SQLiteConnection connection, SQLiteTransaction transaction, int subsectionId)
        {
            var childIds = new List<int>();

            const string childQuery = @"SELECT id FROM Subsection WHERE ParentSubsection_id = @ParentSubsectionId";
            using (var command = new SQLiteCommand(childQuery, connection, transaction))
            {
                command.Parameters.AddWithValue("@ParentSubsectionId", subsectionId);
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        childIds.Add(reader.GetInt32(0));
                    }
                }
            }

            foreach (int childId in childIds)
            {
                await DeleteSubsectionInternalAsync(connection, transaction, childId);
            }

            const string deleteQuery = @"DELETE FROM Subsection WHERE id = @Id";
            using (var command = new SQLiteCommand(deleteQuery, connection, transaction))
            {
                command.Parameters.AddWithValue("@Id", subsectionId);
                await command.ExecuteNonQueryAsync();
            }
        }

        // -----------------------------------------------------------------------------
        // Работа с разделами
        public async Task<int> AddSection(Section section)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO Section (Report_id, Number, Title) 
                        VALUES (@ReportId, @Number, @Title);
                        SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ReportId", section.ReportId);
                    command.Parameters.AddWithValue("@Number", section.Number);
                    command.Parameters.AddWithValue("@Title", section.Title ?? (object)DBNull.Value);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateSection(Section section)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE Section 
                        SET Number = @Number, Title = @Title 
                        WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", section.Id);
                    command.Parameters.AddWithValue("@Number", section.Number);
                    command.Parameters.AddWithValue("@Title", section.Title ?? (object)DBNull.Value);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteSection(int sectionId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM Section WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", sectionId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        // Включение внешнего ключа(не трогать)
        private void EnableForeignKeys(SQLiteConnection connection)
        {
            using (var command = new SQLiteCommand("PRAGMA foreign_keys = ON;", connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }
}
