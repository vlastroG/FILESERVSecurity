using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DirsAccessFromExcel.DTO;
using DirsAccessFromExcel.Logic;
using DirsAccessFromExcele.DTO;
using Microsoft.Office.Interop.Excel;
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

        private readonly List<List<UserDirAccessDto>> ListOfUsersAccRules
            = new List<List<UserDirAccessDto>>();

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

        private void FillRulesList(string @rootDir)
        {
            FillUsernames();
            foreach (var user in Users)
            {
                List<UserDirAccessDto> userAccRules = new List<UserDirAccessDto>();
                var row = 4;
                var columnNumber = 1;
                var columnUser = user.ColumnNumber;
                var path = ReadCell(row, columnNumber);
                var accRule = ReadCell(row, columnUser);
                while ((path != String.Empty) || (accRule != String.Empty))
                {
                    if (path != String.Empty && accRule != string.Empty)
                    {
                        path = rootDir + path;
                        var rule = AccessSetter.GetAccessRule(accRule);
                        userAccRules.Add(new UserDirAccessDto(path, user.Name, rule));
                    }
                    row++;
                    path = ReadCell(row, columnNumber);
                    accRule = ReadCell(row, columnUser);
                }
                ListOfUsersAccRules.Add(userAccRules);
            }
        }

        public void SetAccRulesToDirs(string @root)
        {
            FillRulesList(@root);
            foreach (var accListForUser in ListOfUsersAccRules)
            {
                for (int i = 0; i < accListForUser.Count; i++)
                {
                    AccessSetter.SetDirAccessForUser(accListForUser[i]);
                }
                Console.WriteLine();
            }
        }

        public void Dispose()
        {
            wb.Close(0);
            excel.Quit();
        }
    }
}
