using DirsAccessFromExcele.DTO;
using DirsAccessFromExcele.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcele.Logic
{
    public enum UserAccessRule
    {
        Read,
        Write,
        Disable
    }

    public static class AccessSettter
    {
        public static void SetDirAccessForUser(UserDirAccessDto userDirAccess)
        {
            try
            {
                UserAccessRule rule = userDirAccess.Access;

                switch (rule)
                {
                    case UserAccessRule.Read:
                        SetReadRules(userDirAccess.DirPath, userDirAccess.UserName, userDirAccess.IsDeepAccess);
                        break;
                    case UserAccessRule.Write:
                        SetWriteRules(userDirAccess.DirPath, userDirAccess.UserName, userDirAccess.IsDeepAccess);
                        break;
                    case UserAccessRule.Disable:
                        RemoveAccess(userDirAccess.DirPath, userDirAccess.UserName);
                        break;
                }

                Console.Write($"Назначение доступа для {userDirAccess.UserName} к {userDirAccess.DirPath}");

                ConsoleExtensions.ClearCurrentConsoleLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static void RemoveAccess(string DirPath, string @UserName)
        {
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, true);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, true);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, true);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, true);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, true);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, true);
        }

        private static void SetWriteRules(string DirPath, string @UserName, bool isDeepAccess)
        {
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, isDeepAccess);
        }

        private static void SetReadRules(string DirPath, string @UserName, bool isDeepAccess)
        {
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, isDeepAccess);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, isDeepAccess);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, isDeepAccess);
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, isDeepAccess);
        }

        private static void AddDirectorySecurity(
            string DirPath,
            string UserName,
            FileSystemRights Rights,
            bool isDeepAccess)
        {
            DirectoryInfo dInfo = new DirectoryInfo(DirPath);

            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            // Если применять только для этой папки, то только первое назначение.
            // Если применять для этой папки, ее подпапок и файлов, то нужны и первое и второе назначения.
            dSecurity.AddAccessRule(new FileSystemAccessRule(UserName,
                                                            Rights,
                                                            AccessControlType.Allow)); // первое назначение
            if (isDeepAccess)
            {
                dSecurity.AddAccessRule(new FileSystemAccessRule(UserName,
                                                                Rights,
                                                                InheritanceFlags.ContainerInherit |
                                                                InheritanceFlags.ObjectInherit,
                                                                PropagationFlags.InheritOnly,
                                                                AccessControlType.Allow)); // второе назначение
            }

            dInfo.SetAccessControl(dSecurity);
        }

        // Removes an ACL entry on the specified directory for the specified account.
        private static void RemoveDirectorySecurity(string DirPath, string UserName, FileSystemRights Rights, bool isDeepAccess)
        {
            DirectoryInfo dInfo = new DirectoryInfo(DirPath);

            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            dSecurity.RemoveAccessRule(new FileSystemAccessRule(UserName,
                                                            Rights,
                                                            AccessControlType.Allow));
            if (isDeepAccess)
            {
                dSecurity.RemoveAccessRule(new FileSystemAccessRule(UserName,
                                                Rights,
                                                InheritanceFlags.ContainerInherit |
                                                InheritanceFlags.ObjectInherit,
                                                PropagationFlags.InheritOnly,
                                                AccessControlType.Allow));
            }

            dInfo.SetAccessControl(dSecurity);
        }
    }
}
