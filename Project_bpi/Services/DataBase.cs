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
	                ""SubSectopn_id""	INTEGER NOT NULL,
	                PRIMARY KEY(""Id"" AUTOINCREMENT),
	                FOREIGN KEY(""SubSectopn_id"") REFERENCES ""Template_Subsection""(""Id"") on delete cascade
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
        public Task<ObservableCollection<Template>> GetTemplatesAsync()
        {
            return Task.Run(() =>
            {
                var templates = new ObservableCollection<Template>();
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    EnableForeignKeys(connection);
                    string query = "SELECT Id, Name, Year, Path FROM Templates";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                templates.Add(new Template
                                {
                                    Id = reader.GetInt32(0),
                                    Name = reader.GetString(1),
                                    Year = reader.GetInt32(2),
                                    Path = reader.GetString(3)
                                });
                            }
                            
                        }
                    }
                }
                return templates;
            });
        }
        public Task<ObservableCollection<Report>> GetReportsAsync()
        {
            return Task.Run(() =>
            {
                var reports = new ObservableCollection<Report>();
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    EnableForeignKeys(connection);
                    string query = "SELECT Id, Title, Template_id, Created_at FROM Reports";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                reports.Add(new Report
                                {
                                    Id = reader.GetInt32(0),
                                    Title = reader.GetString(1),
                                    Template_Id = reader.GetInt32(2),
                                    Created_at = reader.GetDateTime(3).ToString()
                                });
                            }
                        }
                    }
                }
                return reports;
            });
        }
        public Task<ObservableCollection<Models.TemplateContent>> GetTemplateContentsAsync(int SubSectionId)
        {
            return Task.Run(() =>
            {
                var contents = new ObservableCollection<Models.TemplateContent>();
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    EnableForeignKeys(connection);
                    string query = $"SELECT Id, Description, Tag, Content_type, SubSectopn_id FROM Template_Content where SubSectionId = {SubSectionId}";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                contents.Add(new Models.TemplateContent
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
                return contents;
            });
        }
        public Task<ObservableCollection<TemplateSection>> GetTemplateSectionsAsync(int TemplateId)
        {
            return Task.Run(() =>
            {
                var sections = new ObservableCollection<TemplateSection>();
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    EnableForeignKeys(connection);
                    string query = $"SELECT Id, Section_number, Title, Template_id FROM Template_Section where Template_id = {TemplateId}";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                sections.Add(new TemplateSection
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
                return sections;
            });
        }
        public Task<ObservableCollection<TemplateSubsection>> GetTemplateSubsectionsAsync(int SectionId)
        {
            return Task.Run(() =>
            {
                var subsections = new ObservableCollection<TemplateSubsection>();
                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    EnableForeignKeys(connection);
                    string query = $"SELECT Id, Subsection_number, Title, Section_id FROM Template_Subsection where Section_id = {SectionId}";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                subsections.Add(new TemplateSubsection
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
                return subsections;
            });
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
