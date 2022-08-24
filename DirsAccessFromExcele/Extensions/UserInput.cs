using System;

namespace DirsAccessFromExcel.Extensions
{
    public static class UserInput
    {
        public static string GetStringFromUser(string MessageToUser)
        {
            Console.WriteLine(MessageToUser);
            Console.Write("-->");
            return Console.ReadLine();
        }
    }
}
