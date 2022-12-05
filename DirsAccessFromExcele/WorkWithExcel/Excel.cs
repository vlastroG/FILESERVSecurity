using DirsAccessFromExcel.DTO;
using DirsAccessFromExcel.Logic;
using DirsAccessFromExcele.DTO;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DirsAccessFromExcel.WorkWithExcel
{
    public class Excel : IDisposable
    {
        public string Path { get; private set; }
        private readonly _Application excel = new _Excel.Application();
        private readonly Workbook wb;
        private readonly Worksheet ws;

        private UsernameDto[] Users;

        /// <summary>
        /// Список Dto скомпонованных абсолютных путей к папкам и прав доступа для пользователей к ним
        /// </summary>
        private readonly List<UserDirAccessDto> ListOfUsersAccRules
            = new List<UserDirAccessDto>();

        /// <summary>
        /// Файл Excel с матрицей прав доступа
        /// </summary>
        /// <param name="Path"></param>
        public Excel(string Path)
        {
            this.Path = Path;
            wb = excel.Workbooks.Open(Path);
            ws = wb.Worksheets[1];
        }

        /// <summary>
        /// Возвращает значение ячейки по номеру строки и столбца. Индексация начинается с 1!
        /// Допускается использовать вместо индекса столбца его название.
        /// </summary>
        /// <param name="row">Номер строки</param>
        /// <param name="column">Номер/название столбца</param>
        /// <returns>Строковое значение ячейки</returns>
        public string ReadCell(int row, object column)
        {
            return ws.Cells[row, column].Value2 ?? String.Empty;
        }

        /// <summary>
        /// Заполняет массив <see cref="Users">Users</see> валидными dto имен пользователей
        /// </summary>
        private void FillUsernames()
        {
            var column = 2;
            var row = 3;
            List<UsernameDto> users = new List<UsernameDto>();
            var cellValue = ReadCell(row, column);
            while (cellValue != String.Empty)
            {
                string[] usersInCell = AccessSetter.GetUsernamesFromString(cellValue);
                foreach (string user in usersInCell)
                {
                    users.Add(new UsernameDto(user, column));
                }
                column++;
                cellValue = ReadCell(row, column);
            }
            Users = users.ToArray();
        }

        /// <summary>
        /// Заполняет сисок <see cref="ListOfUsersAccRules">ListOfUsersAccRules</see>
        /// </summary>
        /// <param name="rootDir">Корневая папка проекта</param>
        private void FillRulesList(string @rootDir)
        {
            Console.WriteLine();
            Console.WriteLine("Подождите, идет обработка матрицы доступа...");
            int row = 4;
            int columnDir = 1;
            int columnUser = 2;
            string path = ReadCell(row, columnDir);
            while (!string.IsNullOrEmpty(path))
            {
                string[] users = AccessSetter.GetUsernamesFromString(ReadCell(3, columnUser));
                while (users.Length > 0)
                {
                    string accRule = ReadCell(row, columnUser);
                    if (!string.IsNullOrEmpty(accRule))
                    {
                        var rule = AccessSetter.GetAccessRule(accRule);
                        foreach (string user in users)
                        {
                            ListOfUsersAccRules.Add(new UserDirAccessDto(rootDir + path, user, rule));
                        }
                    }
                    columnUser++;
                    users = AccessSetter.GetUsernamesFromString(ReadCell(3, columnUser));
                }
                row++;
                columnUser = 2;
                path = ReadCell(row, columnDir);
            }
        }

        /// <summary>
        /// Назначает права доступа к директориям из <see cref="ListOfUsersAccRules">ListOfUsersAccRules</see>
        /// </summary>
        /// <param name="root"></param>
        /// <returns></returns>
        public bool SetAccRulesToDirs(string @root)
        {
            FillRulesList(@root);
            bool result = true;
            foreach (var accListForUser in ListOfUsersAccRules)
            {
                result &= AccessSetter.SetDirAccessForUser(accListForUser);
            }
            return result;
        }

        public void Dispose()
        {
            wb.Close(0);
            excel.Quit();
        }
    }
}
