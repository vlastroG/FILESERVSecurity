using DirsAccessFromExcel.DTO;
using DirsAccessFromExcel.Extensions;
using DirsAccessFromExcel.Logic;
using DirsAccessFromExcel.WorkWithExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            Execute();
        }

        public static void Execute()
        {
            Excel excel = new Excel(@"ExcelSample\ПраваДоступаОбразец.xlsx");
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
                excel.Dispose();

                Console.WriteLine("Готово, нажмите Enter");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                excel.Dispose();
            }

            Console.ReadLine();
        }
    }
}
