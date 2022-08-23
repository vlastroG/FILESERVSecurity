using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcele.Extensions
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
