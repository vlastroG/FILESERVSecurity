using System;

namespace DirsAccessFromExcel.Extensions
{
    /// <summary>
    /// Обработка ввода пользователя
    /// </summary>
    public static class UserInput
    {
        /// <summary>
        /// Получить строку от пользователя без преобразований
        /// </summary>
        /// <param name="MessageToUser"></param>
        /// <returns></returns>
        public static string GetStringFromUser(string MessageToUser)
        {
            Console.WriteLine(MessageToUser);
            Console.Write("-->");
            return Console.ReadLine();
        }
    }
}
