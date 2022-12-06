namespace DirsAccessFromExcel.Logic
{
    /// <summary>
    /// Перечисление прав доступа к папке
    /// </summary>
    public enum UserAccessRule
    {
        /// <summary>
        /// Доступ на чтение только текущей папки
        /// </summary>
        Read,
        /// <summary>
        /// Доступ на чтение текущей папки, ее подпапок и файлов
        /// </summary>
        ReadDeep,
        /// <summary>
        /// Доступ на редактирование текущей папки, ее подпапок и файлов
        /// </summary>
        Write,
        /// <summary>
        /// Доступа нет
        /// </summary>
        Disable
    }
}
