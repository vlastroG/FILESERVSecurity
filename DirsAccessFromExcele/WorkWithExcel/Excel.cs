﻿using DirsAccessFromExcel.DTO;
using DirsAccessFromExcel.Logic;
using DirsAccessFromExcele.DTO;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DirsAccessFromExcel.WorkWithExcel
{
    /// <summary>
    /// Обработчик Excel файла с матрицей доступа
    /// </summary>
    public class Excel : IDisposable
    {
        /// <summary>
        /// Приложение Excel
        /// </summary>
        private readonly _Application excel = new _Excel.Application();

        /// <summary>
        /// Книга Excel с матрицей доступа
        /// </summary>
        private readonly Workbook wb;

        /// <summary>
        /// Лист Excel с матрицей доступа
        /// </summary>
        private readonly Worksheet ws;

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
        /// Полный путь к Excel файлу
        /// </summary>
        public string Path { get; private set; }


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
        /// Назначает права доступа к директориям из <see cref="ListOfUsersAccRules">ListOfUsersAccRules</see>
        /// </summary>
        /// <param name="root"></param>
        /// <returns></returns>
        public bool SetAccRulesToDirs(string @root)
        {
            FillRulesList(@root);
            Console.WriteLine();
            Console.WriteLine(new string('=', 140));
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


        /// <summary>
        /// Заполняет сисок <see cref="ListOfUsersAccRules">ListOfUsersAccRules</see>
        /// </summary>
        /// <param name="rootDir">Корневая папка проекта</param>
        private void FillRulesList(string @rootDir)
        {
            Console.WriteLine();
            Console.WriteLine("Подождите, идет обработка матрицы доступа...");
            FillRulesForRoot(rootDir);
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
        /// Назначает доступ для чтения только корневой папки проекта для всех пользователей из Excel
        /// </summary>
        /// <param name="root">Полный путь к корню проекта</param>
        private void FillRulesForRoot(string @root)
        {
            int columnUser = 2;
            string[] users = AccessSetter.GetUsernamesFromString(ReadCell(3, columnUser));
            while (users.Length > 0)
            {
                foreach (string user in users)
                {
                    ListOfUsersAccRules.Add(new UserDirAccessDto(@root, user, UserAccessRule.Read));
                }
                columnUser++;
                users = AccessSetter.GetUsernamesFromString(ReadCell(3, columnUser));
            }
        }
    }
}
