using DirsAccessFromExcel.DTO;
using DirsAccessFromExcel.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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

    public static class AccessSetter
    {
        public static void SetDirAccessForUser(UserDirAccessDto userDirAccess)
        {
            try
            {
                UserAccessRule rule = userDirAccess.Access;
                Console.WriteLine($"Назначение '{userDirAccess.Access}' доступа для {userDirAccess.UserName} к {userDirAccess.DirPath}");

                switch (rule)
                {
                    case UserAccessRule.Read:
                        SetReadRules(userDirAccess.DirPath, userDirAccess.UserName, false);
                        break;
                    case UserAccessRule.ReadDeep:
                        SetReadRules(userDirAccess.DirPath, userDirAccess.UserName, true);
                        break;
                    case UserAccessRule.Write:
                        SetWriteRules(userDirAccess.DirPath, userDirAccess.UserName);
                        break;
                    case UserAccessRule.Disable:
                        RemoveAccess(userDirAccess.DirPath, userDirAccess.UserName);
                        break;
                }
            }
            catch (IdentityNotMappedException)
            {
                Console.WriteLine($"Ошибка: имя пользователя {userDirAccess.UserName} некорректно !");
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

        private static void SetWriteRules(string DirPath, string @UserName)
        {
            RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, true);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, true);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, true);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, true);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, true);
            AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, true);
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

        private static void DirectoryNotExistMessage(string DirPath)
        {
            Console.WriteLine($"Ошибка: папки {DirPath} не существует !");
        }

        /// <summary>
        /// Возвращает массив имен пользователей из ячейки. В массив попадают только валидные имена.
        /// </summary>
        /// <param name="s">Строковое значение ячейки</param>
        /// <returns>Массив имен пользователей</returns>
        public static string[] GetUsernamesFromString(string s)
        {
            Regex regex = new Regex(@"[a-z]+[.][a-z]{2}$");
            var usersRaw = s.Split(';');
            List<string> users = new List<string>();
            foreach (var user in usersRaw)
            {
                var u = user.StartsWith("\n") ? user.Remove(0, 1) : user;
                if (regex.IsMatch(u))
                    users.Add(u);
            }
            return users.ToArray();
        }

        private static void AddDirectorySecurity(
            string DirPath,
            string UserName,
            FileSystemRights Rights,
            bool isDeepAccess)
        {
            DirectoryInfo dInfo = new DirectoryInfo(DirPath);
            if (!dInfo.Exists)
            {
                DirectoryNotExistMessage(DirPath);
                return;
            }

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
            if (!dInfo.Exists)
            {
                DirectoryNotExistMessage(DirPath);
                return;
            }

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

        /// <summary>
        /// Преобразовывает строку в константу перечисления прав доступа
        /// </summary>
        /// <param name="s">Конвертируемая строка</param>
        /// <returns>Права на редактирование, если "Р";
        /// права на чтение, если "Ч";
        /// права на чтение текущей папки, ее подпапок и файлов, если "Ч!";
        /// иначе доступ запрещен</returns>
        public static UserAccessRule GetAccessRule(string s)
        {
            switch (s)
            {
                case "Р":
                    return UserAccessRule.Write;
                case "Ч":
                    return UserAccessRule.Read;
                case "Ч!":
                    return UserAccessRule.ReadDeep;
                default:
                    return UserAccessRule.Disable;
            }
        }

        /// <summary>
        /// Обнулить доступ для пользователя к проекту
        /// </summary>
        /// <param name="root">Корневая папка проекта</param>
        /// <param name="userNameRaw">Имя пользователя</param>
        public static void RemoveUserAccess(DirectoryInfo root, string userNameRaw)
        {
            string user = GetUsernamesFromString(userNameRaw).FirstOrDefault();
            if (root.Exists)
            {
                string[] subDirs = root.GetSubDirsRecursively().Reverse().ToArray();
                for (int i = 0; i < subDirs.Length; i++)
                {
                    SetDirAccessForUser(new UserDirAccessDto(subDirs[i], user, UserAccessRule.Disable));
                }
            }
            else
            {
                DirectoryNotExistMessage(root.FullName);
            }
        }
    }
}
