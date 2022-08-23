using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcele.DTO
{
    public class UsernameDto
    {
        /// <summary>
        /// Имя пользователя (PGS.ru\ivanov.aa)
        /// </summary>
        public string @Name { get; private set; }

        /// <summary>
        /// Номер столбца, в котором написано имя пользователя, индексация столбцов начинается с 1.
        /// </summary>
        public int ColumnNumber { get; private set; }

        public UsernameDto(string name, int columnNumber)
        {
            @Name = @"PGS.ru\" + name;
            ColumnNumber = columnNumber;
        }
    }
}
