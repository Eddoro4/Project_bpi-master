using Project_bpi.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.IO;
using System.Runtime.InteropServices;
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

        // Создать пустой Word документ по указанному пути (требует установленный MS Word)
        public void CreateEmptyWordDocument(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new InvalidOperationException("Microsoft Word is not installed on this machine.");

            object wordApp = null;
            object document = null;
            try
            {
                wordApp = Activator.CreateInstance(wordType);
                // use dynamic to call COM methods
                dynamic app = wordApp;
                app.Visible = false;
                document = app.Documents.Add();
                app.ActiveDocument.SaveAs2(path);
                app.ActiveDocument.Close();
                app.Quit();
            }
            finally
            {
                if (document != null)
                {
                    try { Marshal.ReleaseComObject(document); } catch { }
                }
                if (wordApp != null)
                {
                    try { Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
        }

        // Загрузка иерархии для одного шаблона
        public Template GetTemplateWithHierarchy(int templateId)
        {
            var template = GetTemplateById(templateId);
            if (template == null) return null;

            var sections = GetSectionsByTemplate(templateId);
            foreach (var section in sections)
            {
                var subsections = GetSubsectionsBySection(section.Id);
                section.Subsections = new List<TemplateSubsection>();
                foreach (var sub in subsections)
                {
                    var contents = GetContentsBySubsection(sub.Id);
                    sub.TemplateContents = contents;
                    section.Subsections.Add(sub);
                }
            }
            template.Sections = sections;
            return template;
        }

        // Загрузка иерархии для всех шаблонов
        public List<Template> GetAllTemplatesWithHierarchy()
        {
            var templates = GetTemplates();
            foreach (var template in templates)
            {
                template.Sections = GetSectionsByTemplate(template.Id);
                foreach (var section in template.Sections)
                {
                    section.Subsections = GetSubsectionsBySection(section.Id);
                    foreach (var sub in section.Subsections)
                    {
                        sub.TemplateContents = GetContentsBySubsection(sub.Id);
                    }
                }
            }
            return templates;
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

                EnableForeignKeys(connection);
                string createTablesScript = @"
                CREATE TABLE IF NOT EXISTS ""Report_data"" (
	            ""Id""	INTEGER,
	            ""Report_id""	INTEGER NOT NULL,
	            ""Content_id""	INTEGER NOT NULL,
	            ""Value""	TEXT,
	            PRIMARY KEY(""Id"" AUTOINCREMENT),
	            FOREIGN KEY(""Content_id"") REFERENCES ""Template_Content""(""Id"") on delete cascade,
	            FOREIGN KEY(""Report_id"") REFERENCES ""Reports""(""Id"") on delete cascade
                );
                CREATE TABLE IF NOT EXISTS ""Reports"" (
	                ""Id""	INTEGER,
	                ""Title""	TEXT,
	                ""Template_id""	INTEGER NOT NULL,
	                ""Created_at""	TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
	                PRIMARY KEY(""Id"" AUTOINCREMENT),
	                FOREIGN KEY(""Template_id"") REFERENCES ""Templates""(""Id"") on delete cascade
                );
                CREATE TABLE IF NOT EXISTS ""Template_Content"" (
	                ""Id""	INTEGER,
	                ""Description""	TEXT,
	                ""Tag""	TEXT NOT NULL,
	                ""Content_type""	TEXT NOT NULL CHECK(""Content_type"" IN ('text', 'table')),
	                ""SubSection_id""	INTEGER NOT NULL,
	                PRIMARY KEY(""Id"" AUTOINCREMENT),
	                FOREIGN KEY(""SubSection_id"") REFERENCES ""Template_Subsection""(""Id"") on delete cascade
                );
                CREATE TABLE IF NOT EXISTS ""Template_Section"" (
	                ""Id""	INTEGER,
                    ""Section_number""	INTEGER NOT NULL,
	                ""Title""	TEXT NOT NULL,
	                ""Template_id""	INTEGER NOT NULL,
	                PRIMARY KEY(""Id"" AUTOINCREMENT),
	                FOREIGN KEY(""Template_id"") REFERENCES ""Templates""(""Id"") on delete cascade
                );
                CREATE TABLE IF NOT EXISTS ""Template_Subsection"" (
	                ""Id""	INTEGER,
                    ""Subsection_number""	INTEGER NOT NULL,
	                ""Title""	TEXT NOT NULL,
	                ""Section_id""	INTEGER NOT NULL,
	                PRIMARY KEY(""Id"" AUTOINCREMENT),
	                FOREIGN KEY(""Section_id"") REFERENCES ""Template_Section""(""Id"") on delete cascade
                );
                CREATE TABLE IF NOT EXISTS ""Templates"" (
	                ""Id""	INTEGER,
	                ""Name""	TEXT,
	                ""Year""	INTEGER NOT NULL,
                    ""Path"" TEXT NOT NULL,
	                PRIMARY KEY(""Id"" AUTOINCREMENT)
                );";
                using (var command = new SQLiteCommand(createTablesScript, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
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

        // --------------------- Template methods ---------------------
        public List<Template> GetTemplates()
        {
            var result = new List<Template>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var command = new SQLiteCommand("SELECT Id, Name, Year, Path FROM Templates", connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new Template
                        {
                            Id = reader.GetInt32(0),
                            Name = reader.IsDBNull(1) ? null : reader.GetString(1),
                            Year = reader.GetInt32(2),
                            Path = reader.IsDBNull(3) ? null : reader.GetString(3)
                        });
                    }
                }
            }
            return result;
        }

        public Template GetTemplateById(int id)
        {
            Template template = null;
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Name, Year, Path FROM Templates WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            template = new Template
                            {
                                Id = reader.GetInt32(0),
                                Name = reader.IsDBNull(1) ? null : reader.GetString(1),
                                Year = reader.GetInt32(2),
                                Path = reader.IsDBNull(3) ? null : reader.GetString(3)
                            };
                        }
                    }
                }
            }
            return template;
        }

        public int InsertTemplate(Template template)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Templates (Name, Year, Path) VALUES (@name, @year, @path);", connection))
                {
                    cmd.Parameters.AddWithValue("@name", (object)template.Name ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@year", template.Year);
                    cmd.Parameters.AddWithValue("@path", template.Path ?? string.Empty);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void UpdateTemplate(Template template)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("UPDATE Templates SET Name = @name, Year = @year, Path = @path WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@name", (object)template.Name ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@year", template.Year);
                    cmd.Parameters.AddWithValue("@path", template.Path ?? string.Empty);
                    cmd.Parameters.AddWithValue("@id", template.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteTemplate(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Templates WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // --------------------- Sections / Subsections / Contents ---------------------
        public List<TemplateSection> GetSectionsByTemplate(int templateId)
        {
            var result = new List<TemplateSection>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Section_number, Title, Template_id FROM Template_Section WHERE Template_id = @tid ORDER BY Section_number", connection))
                {
                    cmd.Parameters.AddWithValue("@tid", templateId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new TemplateSection
                            {
                                Id = reader.GetInt32(0),
                                Section_Number = reader.GetInt32(1),
                                Title = reader.GetString(2),
                                Template_Id = reader.GetInt32(3)
                            });
                        }
                    }
                }
            }
            return result;
        }

        public int InsertSection(TemplateSection section)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Template_Section (Section_number, Title, Template_id) VALUES (@num, @title, @tid)", connection))
                {
                    cmd.Parameters.AddWithValue("@num", section.Section_Number);
                    cmd.Parameters.AddWithValue("@title", section.Title ?? string.Empty);
                    cmd.Parameters.AddWithValue("@tid", section.Template_Id);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void UpdateSection(TemplateSection section)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("UPDATE Template_Section SET Section_number = @num, Title = @title WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@num", section.Section_Number);
                    cmd.Parameters.AddWithValue("@title", section.Title ?? string.Empty);
                    cmd.Parameters.AddWithValue("@id", section.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteSection(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Template_Section WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public List<TemplateSubsection> GetSubsectionsBySection(int sectionId)
        {
            var result = new List<TemplateSubsection>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Subsection_number, Title, Section_id FROM Template_Subsection WHERE Section_id = @sid ORDER BY Subsection_number", connection))
                {
                    cmd.Parameters.AddWithValue("@sid", sectionId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new TemplateSubsection
                            {
                                Id = reader.GetInt32(0),
                                Subsection_Number = reader.GetInt32(1),
                                Title = reader.GetString(2),
                                SectionId = reader.GetInt32(3)
                            });
                        }
                    }
                }
            }
            return result;
        }

        public int InsertSubsection(TemplateSubsection subsection)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Template_Subsection (Subsection_number, Title, Section_id) VALUES (@num, @title, @sid)", connection))
                {
                    cmd.Parameters.AddWithValue("@num", subsection.Subsection_Number);
                    cmd.Parameters.AddWithValue("@title", subsection.Title ?? string.Empty);
                    cmd.Parameters.AddWithValue("@sid", subsection.SectionId);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void UpdateSubsection(TemplateSubsection subsection)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("UPDATE Template_Subsection SET Subsection_number = @num, Title = @title WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@num", subsection.Subsection_Number);
                    cmd.Parameters.AddWithValue("@title", subsection.Title ?? string.Empty);
                    cmd.Parameters.AddWithValue("@id", subsection.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteSubsection(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Template_Subsection WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public List<Models.TemplateContent> GetContentsBySubsection(int subsectionId)
        {
            var result = new List<Models.TemplateContent>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Description, Tag, Content_type, SubSection_id FROM Template_Content WHERE SubSection_id = @ssid ORDER BY Id", connection))
                {
                    cmd.Parameters.AddWithValue("@ssid", subsectionId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new Models.TemplateContent
                            {
                                Id = reader.GetInt32(0),
                                Description = reader.IsDBNull(1) ? null : reader.GetString(1),
                                Tag = reader.GetString(2),
                                ContentType = reader.GetString(3),
                                Subsection_Id = reader.GetInt32(4)
                            });
                        }
                    }
                }
            }
            return result;
        }

        public int InsertContent(Models.TemplateContent content)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Template_Content (Description, Tag, Content_type, SubSection_id) VALUES (@desc, @tag, @ctype, @ssid)", connection))
                {
                    cmd.Parameters.AddWithValue("@desc", (object)content.Description ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@tag", content.Tag ?? string.Empty);
                    cmd.Parameters.AddWithValue("@ctype", content.ContentType ?? string.Empty);
                    cmd.Parameters.AddWithValue("@ssid", content.Subsection_Id);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void UpdateContent(Models.TemplateContent content)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("UPDATE Template_Content SET Description = @desc, Tag = @tag, Content_type = @ctype WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@desc", (object)content.Description ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@tag", content.Tag ?? string.Empty);
                    cmd.Parameters.AddWithValue("@ctype", content.ContentType ?? string.Empty);
                    cmd.Parameters.AddWithValue("@id", content.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteContent(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Template_Content WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // --------------------- Reports & ReportData ---------------------
        public List<Report> GetReports()
        {
            var result = new List<Report>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Title, Template_id, Created_at FROM Reports ORDER BY Created_at DESC", connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new Report
                        {
                            Id = reader.GetInt32(0),
                            Title = reader.IsDBNull(1) ? null : reader.GetString(1),
                            Template_Id = reader.GetInt32(2),
                            Created_at = reader.IsDBNull(3) ? null : reader.GetString(3)
                        });
                    }
                }
            }
            return result;
        }

        public int InsertReport(Report report)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Reports (Title, Template_id) VALUES (@title, @tid)", connection))
                {
                    cmd.Parameters.AddWithValue("@title", (object)report.Title ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@tid", report.Template_Id);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void DeleteReport(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Reports WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public List<ReportData> GetReportDataByReportId(int reportId)
        {
            var result = new List<ReportData>();
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("SELECT Id, Report_id, Content_id, Value FROM Report_data WHERE Report_id = @rid", connection))
                {
                    cmd.Parameters.AddWithValue("@rid", reportId);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new ReportData
                            {
                                Id = reader.GetInt32(0),
                                Report_Id = reader.GetInt32(1),
                                Content_Id = reader.GetInt32(2),
                                Value = reader.IsDBNull(3) ? null : reader.GetString(3)
                            });
                        }
                    }
                }
            }
            return result;
        }

        public int InsertReportData(ReportData data)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("INSERT INTO Report_data (Report_id, Content_id, Value) VALUES (@rid, @cid, @val)", connection))
                {
                    cmd.Parameters.AddWithValue("@rid", data.Report_Id);
                    cmd.Parameters.AddWithValue("@cid", data.Content_Id);
                    cmd.Parameters.AddWithValue("@val", (object)data.Value ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
                using (var last = new SQLiteCommand("SELECT last_insert_rowid();", connection))
                {
                    return Convert.ToInt32(last.ExecuteScalar());
                }
            }
        }

        public void UpdateReportData(ReportData data)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("UPDATE Report_data SET Value = @val WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@val", (object)data.Value ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@id", data.Id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void DeleteReportData(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                EnableForeignKeys(connection);
                using (var cmd = new SQLiteCommand("DELETE FROM Report_data WHERE Id = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // Закрытие соединения с бд
        public void Dispose()
        {
            _connection?.Dispose();
        }
    }
}
