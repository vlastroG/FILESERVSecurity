using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcele
{
    public class Program
    {
        static void Main(string[] args)
        {
            Execute();
        }

        public static void Execute()
        {
            try
            {
                string DirectoryName = @"C:\Users\stroganov.vg\Documents\TestAccess_deleteIfYouWant";

                Console.WriteLine("Adding access control entry for " + DirectoryName);

                // Add the access control entry to the directory.
                AddDirectorySecurity(DirectoryName, @"PGS.ru\koval.mi", FileSystemRights.Read, AccessControlType.Allow);
                AddDirectorySecurity(DirectoryName, @"PGS.ru\koval.mi", FileSystemRights.ReadAndExecute, AccessControlType.Allow);
                AddDirectorySecurity(DirectoryName, @"PGS.ru\koval.mi", FileSystemRights.ListDirectory, AccessControlType.Allow);

                //Console.WriteLine("Removing access control entry from " + DirectoryName);

                // Remove the access control entry from the directory.
                //RemoveDirectorySecurity(DirectoryName, @"MYDOMAIN\MyAccount", FileSystemRights.Read, AccessControlType.Allow);

                Console.WriteLine("Done.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            // Console.ReadLine();
        }

        // Adds an ACL entry on the specified directory for the specified account.
        public static void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            dSecurity.AddAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType)); // первое назначение
            dSecurity.AddAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            InheritanceFlags.ContainerInherit |
                                                            InheritanceFlags.ObjectInherit,
                                                            PropagationFlags.InheritOnly,
                                                            ControlType)); // второе назначение
            // Если применять только для этой папки, то только первое назначение.
            // Если применять для этой папки, ее подпапок и файлов, то нужны и первое и второе назначения.

            dInfo.SetAccessControl(dSecurity);
        }

        // Removes an ACL entry on the specified directory for the specified account.
        public static void RemoveDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            dSecurity.RemoveAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType));

            dInfo.SetAccessControl(dSecurity);
        }
    }
}
