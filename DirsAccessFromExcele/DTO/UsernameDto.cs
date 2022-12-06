namespace DirsAccessFromExcele.DTO
{
    /// <summary>
    /// Dto для имен пользователей с доменными префиксами и номеров столбца Excel
    /// </summary>
    public class UsernameDto
    {
        /// <summary>
        /// Имя пользователя (PGS.ru\ivanov.aa)
        /// </summary>
        private readonly string _name;

        /// <summary>
        /// Номер столбца, в котором написано имя пользователя, индексация столбцов начинается с 1.
        /// </summary>
        private readonly int _column;


        /// <summary>
        /// Конструктор Dto имени пользователя и номера столюца из Excel
        /// </summary>
        /// <param name="name">Имя пользователя (ivanov.ii)</param>
        /// <param name="columnNumber">Номер столбца в Excel (начиная с 1)</param>
        public UsernameDto(string name, int columnNumber)
        {
            _name = @"PGS.ru\" + name;
            _column = columnNumber;
        }


        /// <summary>
        /// Имя пользователя с доменным префиксом (PGS.ru\ivanov.aa)
        /// </summary>
        public string @Name { get => _name; }

        /// <summary>
        /// Номер столбца, в котором написано имя пользователя, индексация столбцов начинается с 1.
        /// </summary>
        public int ColumnNumber { get => _column; }
    }
}
