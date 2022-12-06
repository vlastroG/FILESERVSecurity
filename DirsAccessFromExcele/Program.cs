using DirsAccessFromExcel.Extensions;
using DirsAccessFromExcel.Logic;
using DirsAccessFromExcel.WorkWithExcel;
using System;
using System.IO;

namespace DirsAccessFromExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            var input = GetAction();
            while (input != "выход")
            {
                switch (input)
                {
                    case "назначить":
                        var result = SetAccessByExcel();
                        if (result)
                        {
                            Console.WriteLine();
                            Console.WriteLine(new string('=', 140));
                            Console.WriteLine("Готово, нажмите Enter -->");
                        }
                        else
                        {
                            Console.WriteLine();
                            Console.WriteLine(new string('=', 140));
                            Console.WriteLine("Выполнено с ошибками, нажмите Enter -->");
                        }
                        Console.ReadKey(true);
                        Console.Clear();
                        break;
                    case "обнулить":
                        result = DisableAccessForUser();
                        if (result)
                        {
                            Console.WriteLine("Готово, нажмите Enter");
                        }
                        else
                        {
                            Console.WriteLine("Выполнено с ошибками, нажмите Enter");
                        }
                        Console.ReadLine();
                        Console.Clear();
                        break;
                    default:
                        Console.Clear();
                        break;
                }
                input = GetAction();
            }
        }

        /// <summary>
        /// Получить строку команды от пользователя
        /// </summary>
        /// <returns></returns>
        private static string GetAction()
        {
            return UserInput.GetStringFromUser(
                "Если вы хотите назначить права доступа к папкам проекта введите \'назначить\';" +
                "\nесли вы хотите обнулить доступ для пользователя к папкам проекта введите \'обнулить\';" +
                "\nдля выхода введите \'выход\'." +
                "\nРегистр не важен.").TrimStart().TrimEnd().ToLower();
        }

        /// <summary>
        /// Выполнить команду запрещения доступа для пользователя ко всем папкам проекта
        /// </summary>
        /// <returns>True, если команда выполнена успешно, иначе false</returns>
        public static bool DisableAccessForUser()
        {
            string userRaw = UserInput.GetStringFromUser("Введите имя пользователя для запрета доступа," +
                " например \'ivanov.ii\'").TrimStart().TrimEnd();
            string path = UserInput.GetStringFromUser("Введите полный путь к корневой папке проекта,\n" +
                    "например: \'Q:\\Проекты\\003-2022-ПИР\' и нажмите Enter").TrimStart().TrimEnd();
            DirectoryInfo dir = new DirectoryInfo(path);
            return AccessSetter.RemoveUserAccess(dir, userRaw);
        }

        /// <summary>
        /// Выполнить команду назначения прав доступа по таблице Excel
        /// </summary>
        /// <returns></returns>
        public static bool SetAccessByExcel()
        {
            Excel excel = new Excel($"{AppDomain.CurrentDomain.BaseDirectory}ExcelSample\\ПраваДоступаОбразец.xlsx");
            try
            {
                Console.WriteLine();
                string rootDir = UserInput.GetStringFromUser("Введите полный путь к корневой папке проекта,\n" +
                    "например: \'Q:\\Проекты\\003-2022-ПИР\\\' и нажмите Enter").TrimStart().TrimEnd();
                Console.WriteLine();
                string excelPath = UserInput.GetStringFromUser(
                    $"Введите полный путь к excel файлу с правами доступа для {rootDir},\n" +
                    $"например: C:\\Users\\ПраваДоступаОбразец.xlsx");
                Console.WriteLine();
                excel = new Excel(excelPath);
                return excel.SetAccRulesToDirs(rootDir);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
            finally
            {
                excel.Dispose();
            }
        }
    }
}
