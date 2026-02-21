using Project_bpi.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Project_bpi.Services
{
    public class DataBase : IDisposable
    {
        private readonly string _connectionString;
        private readonly string databasePath;
        private SQLiteConnection _connection;
        
        public DataBase()
        {
            databasePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Kurs.db");
            _connectionString = $"Data Source={databasePath};";
        }
        // Запуск Базы данных
        public void InitializeDatabase()
        {
            bool databaseExists = File.Exists(databasePath);
            if (!databaseExists)
            {
                File.Create(databasePath).Close();
                MessageBox.Show("Файл базы данных не был обнаружен, создан новый файл базы данных.");
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
                    ""Number"" INTEGER NOT NULL,
                    ""Title"" TEXT,
                    FOREIGN KEY(""Section_id"") REFERENCES ""Section""(""id"") ON DELETE CASCADE
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
                    FOREIGN KEY(""Table_id"") REFERENCES ""Table""(""id"") ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS ""Text"" (
                    ""id"" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                    ""Subsection_id"" INTEGER NOT NULL,
                    ""Content"" TEXT NOT NULL,
                    ""Pattern_name"" TEXT NOT NULL,
                    FOREIGN KEY(""Subsection_id"") REFERENCES ""Subsection""(""id"") ON DELETE CASCADE
                );";
                using (var command = new SQLiteCommand(createTablesScript, connection))
                {
                    command.ExecuteNonQuery();
                }
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

        // Вывод всех шаблонов из бд
        public async Task<List<PatternFile>> GetAllFilePatterns()
        {
            var patterns = new List<PatternFile>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "SELECT id, Title, Year, Path FROM FilePatterns";
                using (var command = new SQLiteCommand(query, connection))
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        patterns.Add(new PatternFile
                        {
                            Id = reader.GetInt32(0),
                            Title = reader.GetString(1),
                            Year = reader.GetInt32(2),
                            Path = reader.GetString(3)
                        });
                    }
                }
            }
            return patterns;
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
                    string sectionsQuery = @"SELECT id, Report_id, Number, Title 
                                        FROM Section WHERE Report_id = @ReportId 
                                        ORDER BY Number";

                    using (var command = new SQLiteCommand(sectionsQuery, connection))
                    {
                        command.Parameters.AddWithValue("@ReportId", reportId);
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                var section = new Section
                                {
                                    Id = reader.GetInt32(0),
                                    ReportId = reader.GetInt32(1),
                                    Number = reader.GetInt32(2),
                                    Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                    SubSections = new List<SubSection>()
                                };
                                report.Sections.Add(section);
                            }
                        }
                    }
                    foreach (var section in report.Sections)
                    {
                        string subsectionsQuery = @"SELECT id, Section_id, Number, Title 
                                               FROM Subsection WHERE Section_id = @SectionId 
                                               ORDER BY Number";

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
                                        Number = reader.GetInt32(2),
                                        Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                        Tables = new List<Table>(),
                                        Texts = new List<Text>()
                                    };
                                    section.SubSections.Add(subsection);
                                }
                            }
                        }
                        foreach (var subsection in section.SubSections)
                        {
                            string tablesQuery = @"SELECT id, Title, Subsection_id, Pattern_name 
                                              FROM Table WHERE Subsection_id = @SubsectionId";

                            using (var command = new SQLiteCommand(tablesQuery, connection))
                            {
                                command.Parameters.AddWithValue("@SubsectionId", subsection.Id);
                                using (var reader = await command.ExecuteReaderAsync())
                                {
                                    while (await reader.ReadAsync())
                                    {
                                        var table = new Table
                                        {
                                            Id = reader.GetInt32(0),
                                            Title = reader.GetString(1),
                                            SubsectionId = reader.GetInt32(2),
                                            PatternName = reader.GetString(3),
                                            TableItems = new List<TableItem>()
                                        };
                                        subsection.Tables.Add(table);
                                    }
                                }
                            }
                            string textsQuery = @"SELECT id, Subsection_id, Text, Pattern_name 
                                             FROM Text WHERE Subsection_id = @SubsectionId";

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
                                string itemsQuery = @"SELECT id, Table_id, Column, Row, Header 
                                                 FROM Table_item WHERE Table_id = @TableId 
                                                 ORDER BY Row, Column";

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
                                                Header = reader.GetString(4)
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return report;
        }

        // Удалить отчёт
        public async Task DeleteReport(int reportId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        string query = "DELETE FROM Report WHERE id = @Id";
                        using (var command = new SQLiteCommand(query, connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", reportId);
                            await command.ExecuteNonQueryAsync();
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
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

                string query = @"INSERT INTO Table (Title, Subsection_id, Pattern_name) 
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

                string query = @"UPDATE Table 
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

                string query = "DELETE FROM Table WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", tableId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task<Table> GetTableWithItems(int tableId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                Table table = null;

                string tableQuery = "SELECT id, Title, Subsection_id, Pattern_name FROM Table WHERE id = @TableId";
                using (var command = new SQLiteCommand(tableQuery, connection))
                {
                    command.Parameters.AddWithValue("@TableId", tableId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        if (await reader.ReadAsync())
                        {
                            table = new Table
                            {
                                Id = reader.GetInt32(0),
                                Title = reader.GetString(1),
                                SubsectionId = reader.GetInt32(2),
                                PatternName = reader.GetString(3),
                                TableItems = new ObservableCollection<TableItem>()
                            };
                        }
                    }
                }

                if (table != null)
                {
                    string itemsQuery = "SELECT id, Table_id, Column, Row, Header FROM Table_item WHERE Table_id = @TableId ORDER BY Row, Column";
                    using (var command = new SQLiteCommand(itemsQuery, connection))
                    {
                        command.Parameters.AddWithValue("@TableId", tableId);
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
                                    Header = reader.GetString(4)
                                });
                            }
                        }
                    }
                }

                return table;
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

                string query = @"INSERT INTO Table_item (Table_id, Column, Row, Header) 
                            VALUES (@TableId, @Column, @Row, @Header);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@TableId", item.TableId);
                    command.Parameters.AddWithValue("@Column", item.Column);
                    command.Parameters.AddWithValue("@Row", item.Row);
                    command.Parameters.AddWithValue("@Header", item.Header);

                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        public async Task UpdateTableItem(TableItem item)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"UPDATE Table_item 
                            SET Column = @Column, Row = @Row, Header = @Header 
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", item.Id);
                    command.Parameters.AddWithValue("@Column", item.Column);
                    command.Parameters.AddWithValue("@Row", item.Row);
                    command.Parameters.AddWithValue("@Header", item.Header);

                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteTableItem(int itemId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM Table_item WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", itemId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task DeleteAllTableItems(int tableId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "DELETE FROM Table_item WHERE Table_id = @TableId";
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

                string query = @"INSERT INTO Text (Subsection_id, Text, Pattern_name) 
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

                string query = @"UPDATE Text 
                            SET Text = @Content, Pattern_name = @PatternName 
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

                string query = "DELETE FROM Text WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", textId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task<List<Text>> GetTextsBySubsection(int subsectionId)
        {
            var texts = new List<Text>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "SELECT id, Subsection_id, Text, Pattern_name FROM Text WHERE Subsection_id = @SubsectionId";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@SubsectionId", subsectionId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            texts.Add(new Text
                            {
                                Id = reader.GetInt32(0),
                                SubsectionId = reader.GetInt32(1),
                                Content = reader.GetString(2),
                                PatternName = reader.GetString(3)
                            });
                        }
                    }
                }
            }

            return texts;
        }
        // -----------------------------------------------------------------------------
        // работа с подразрелами
        public async Task<int> AddSubsection(SubSection subsection)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = @"INSERT INTO Subsection (Section_id, Number, Title) 
                            VALUES (@SectionId, @Number, @Title);
                            SELECT last_insert_rowid();";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@SectionId", subsection.SectionId);
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
                            SET Number = @Number, Title = @Title 
                            WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", subsection.Id);
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

                string query = "DELETE FROM Subsection WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", subsectionId);
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task<SubSection> GetSubsectionWithContents(int subsectionId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                SubSection subsection = null;

                string subsectionQuery = "SELECT id, Section_id, Number, Title FROM Subsection WHERE id = @Id";
                using (var command = new SQLiteCommand(subsectionQuery, connection))
                {
                    command.Parameters.AddWithValue("@Id", subsectionId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        if (await reader.ReadAsync())
                        {
                            subsection = new SubSection
                            {
                                Id = reader.GetInt32(0),
                                SectionId = reader.GetInt32(1),
                                Number = reader.GetInt32(2),
                                Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                Tables = new ObservableCollection<Table>(),
                                Texts = new ObservableCollection<Text>()
                            };
                        }
                    }
                }

                if (subsection != null)
                {
                    // Получаем таблицы
                    string tablesQuery = "SELECT id, Title, Subsection_id, Pattern_name FROM Table WHERE Subsection_id = @SubsectionId";
                    using (var command = new SQLiteCommand(tablesQuery, connection))
                    {
                        command.Parameters.AddWithValue("@SubsectionId", subsectionId);
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                var table = new Table
                                {
                                    Id = reader.GetInt32(0),
                                    Title = reader.GetString(1),
                                    SubsectionId = reader.GetInt32(2),
                                    PatternName = reader.GetString(3),
                                    TableItems = new ObservableCollection<TableItem>()
                                };
                                subsection.Tables.Add(table);
                            }
                        }
                    }

                    // Получаем тексты
                    string textsQuery = "SELECT id, Subsection_id, Text, Pattern_name FROM Text WHERE Subsection_id = @SubsectionId";
                    using (var command = new SQLiteCommand(textsQuery, connection))
                    {
                        command.Parameters.AddWithValue("@SubsectionId", subsectionId);
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

                    // Для каждой таблицы получаем элементы
                    foreach (var table in subsection.Tables)
                    {
                        string itemsQuery = "SELECT id, Table_id, Column, Row, Header FROM Table_item WHERE Table_id = @TableId ORDER BY Row, Column";
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
                                        Header = reader.GetString(4)
                                    });
                                }
                            }
                        }
                    }
                }

                return subsection;
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

        public async Task<Section> GetSectionWithContents(int sectionId)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                Section section = null;

                string sectionQuery = "SELECT id, Report_id, Number, Title FROM Section WHERE id = @Id";
                using (var command = new SQLiteCommand(sectionQuery, connection))
                {
                    command.Parameters.AddWithValue("@Id", sectionId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        if (await reader.ReadAsync())
                        {
                            section = new Section
                            {
                                Id = reader.GetInt32(0),
                                ReportId = reader.GetInt32(1),
                                Number = reader.GetInt32(2),
                                Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                SubSections = new ObservableCollection<SubSection>()
                            };
                        }
                    }
                }

                if (section != null)
                {
                    string subsectionsQuery = @"SELECT id, Section_id, Number, Title 
                                       FROM Subsection 
                                       WHERE Section_id = @SectionId 
                                       ORDER BY Number";

                    using (var command = new SQLiteCommand(subsectionsQuery, connection))
                    {
                        command.Parameters.AddWithValue("@SectionId", sectionId);
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                var subsection = new SubSection
                                {
                                    Id = reader.GetInt32(0),
                                    SectionId = reader.GetInt32(1),
                                    Number = reader.GetInt32(2),
                                    Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                    Tables = new ObservableCollection<Table>(),
                                    Texts = new ObservableCollection<Text>()
                                };
                                section.SubSections.Add(subsection);
                            }
                        }
                    }

                    foreach (var subsection in section.SubSections)
                    {
                        string tablesQuery = "SELECT id, Title, Subsection_id, Pattern_name FROM Table WHERE Subsection_id = @SubsectionId";
                        using (var command = new SQLiteCommand(tablesQuery, connection))
                        {
                            command.Parameters.AddWithValue("@SubsectionId", subsection.Id);
                            using (var reader = await command.ExecuteReaderAsync())
                            {
                                while (await reader.ReadAsync())
                                {
                                    var table = new Table
                                    {
                                        Id = reader.GetInt32(0),
                                        Title = reader.GetString(1),
                                        SubsectionId = reader.GetInt32(2),
                                        PatternName = reader.GetString(3),
                                        TableItems = new ObservableCollection<TableItem>()
                                    };
                                    subsection.Tables.Add(table);
                                }
                            }
                        }
                        string textsQuery = "SELECT id, Subsection_id, Text, Pattern_name FROM Text WHERE Subsection_id = @SubsectionId";
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
                            string itemsQuery = "SELECT id, Table_id, Column, Row, Header FROM Table_item WHERE Table_id = @TableId ORDER BY Row, Column";
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
                                            Header = reader.GetString(4)
                                        });
                                    }
                                }
                            }
                        }
                    }
                }

                return section;
            }
        }

        public async Task<List<Section>> GetSectionsByReport(int reportId)
        {
            var sections = new List<Section>();

            using (var connection = new SQLiteConnection(_connectionString))
            {
                await connection.OpenAsync();
                EnableForeignKeys(connection);

                string query = "SELECT id, Report_id, Number, Title FROM Section WHERE Report_id = @ReportId ORDER BY Number";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ReportId", reportId);
                    using (var reader = await command.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            sections.Add(new Section
                            {
                                Id = reader.GetInt32(0),
                                ReportId = reader.GetInt32(1),
                                Number = reader.GetInt32(2),
                                Title = reader.IsDBNull(3) ? null : reader.GetString(3),
                                SubSections = new ObservableCollection<SubSection>()
                            });
                        }
                    }
                }
            }

            return sections;
        }
        // --------------------------------------------------------------------------------

        // Включение внешнего ключа(не трогать)
        private void EnableForeignKeys(SQLiteConnection connection)
        {
            using (var command = new SQLiteCommand("PRAGMA foreign_keys = ON;", connection))
            {
                command.ExecuteNonQuery();
            }
        }

        // Закрытие соединения с бд
        public void Dispose()
        {
            _connection?.Dispose();
        }
    }
}
