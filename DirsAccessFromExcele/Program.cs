using DirsAccessFromExcele.DTO;
using DirsAccessFromExcele.Extensions;
using DirsAccessFromExcele.Logic;
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
            try
            {
                string DirectoryName = @"C:\Users\stroganov.vg\Documents\TestAccess_deleteIfYouWant";
                string username = @"PGS.ru\koval.mi";
                string rootDir = UserInput.GetStringFromUser("Введите полный путь к корневой папке проекта, например \'Q:\\Проекты\\003-2022-ПИР\'").TrimStart().TrimEnd();
                Console.WriteLine("Test begun");

                UserDirAccessDto dto = new UserDirAccessDto(DirectoryName, username, UserAccessRule.Disable, true);
                AccessSettter.SetDirAccessForUser(dto);

                Console.WriteLine("Done.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.ReadLine();
        }
    }
}
