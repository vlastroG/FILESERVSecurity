using DirsAccessFromExcel.Logic;

namespace DirsAccessFromExcel.DTO
{
    /// <summary>
    /// Dto для компоновки прав доступа к папке для пользователя
    /// </summary>
    public class UserDirAccessDto
    {
        /// <summary>
        /// Путь к папке
        /// </summary>
        private readonly string _path;

        /// <summary>
        /// Имя пользователя (ivanov.ii)
        /// </summary>
        private readonly string _user;

        /// <summary>
        /// Уровень доступа к папке
        /// </summary>
        private readonly UserAccessRule _rule;


        /// <summary>
        /// Конструктор Dto для назначения прав доступа к папке для пользователя
        /// </summary>
        /// <param name="DirPath">Абсолютный путь к папке</param>
        /// <param name="UserName">Имя пользователя (ivanov.ii)</param>
        /// <param name="Access">Уровень доступа к папке</param>
        public UserDirAccessDto(string DirPath, string UserName, UserAccessRule Access)
        {
            _path = DirPath;
            _user = UserName;
            _rule = Access;
        }


        /// <summary>
        /// Абсолютный путь к папке
        /// </summary>
        public string @DirPath { get => _path; }

        /// <summary>
        /// Имя пользователя (ivanov.ii)
        /// </summary>
        public string @UserName { get => _user; }

        /// <summary>
        /// Уровень доступа к папке
        /// </summary>
        public UserAccessRule Access { get => _rule; }
    }
}
