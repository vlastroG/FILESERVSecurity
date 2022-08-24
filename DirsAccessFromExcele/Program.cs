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
                        SetAccessByExcel();
                        Console.WriteLine("Готово, нажмите Enter");
                        Console.ReadLine();
                        Console.Clear();
                        break;
                    case "обнулить":
                        DisableAccessForUser();
                        Console.WriteLine("Готово, нажмите Enter");
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

        private static string GetAction()
        {
            return UserInput.GetStringFromUser(
                "Если вы хотите назначить права доступа к папкам проекта введите \'назначить\';" +
                "\nесли вы хотите обнулить доступ для пользователя к папкам проекта введите \'обнулить\';" +
                "\nдля выхода введите \'выход\'." +
                "\nРегистр не важен.").TrimStart().TrimEnd().ToLower();
        }

        public static void DisableAccessForUser()
        {
            string userRaw = UserInput.GetStringFromUser("Введите имя пользователя для запрета доступа," +
                " например \'ivanov.ii\'").TrimStart().TrimEnd();
            string path = UserInput.GetStringFromUser("Введите полный путь к корневой папке проекта,\n" +
                    "например: \'Q:\\Проекты\\003-2022-ПИР\' и нажмите Enter").TrimStart().TrimEnd();
            DirectoryInfo dir = new DirectoryInfo(path);
            AccessSetter.RemoveUserAccess(dir, userRaw);
        }

        public static void SetAccessByExcel()
        {
            Excel excel = new Excel($"{AppDomain.CurrentDomain.BaseDirectory}ExcelSample\\ПраваДоступаОбразец.xlsx");
            try
            {
                string rootDir = UserInput.GetStringFromUser("Введите полный путь к корневой папке проекта,\n" +
                    "например: \'Q:\\Проекты\\003-2022-ПИР\\\' и нажмите Enter").TrimStart().TrimEnd();
                Console.WriteLine();
                string excelPath = UserInput.GetStringFromUser(
                    $"Введите полный путь к excel файлу с правами доступа для {rootDir},\n" +
                    $"например: C:\\Users\\ПраваДоступаОбразец.xlsx");
                Console.WriteLine();
                excel = new Excel(excelPath);
                excel.SetAccRulesToDirs(rootDir);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                excel.Dispose();
            }
        }
    }
}
