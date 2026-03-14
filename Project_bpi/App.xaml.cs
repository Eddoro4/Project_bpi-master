using Project_bpi.Services;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Project_bpi
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static DataBase DB = new DataBase();
        public App()
        {
            // Инициализация базы данных при запуске приложения
            DB.InitializeDatabase();
        }
    }
}
