﻿using DirsAccessFromExcel.DTO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.RegularExpressions;

namespace DirsAccessFromExcel.Logic
{
    /// <summary>
    /// Управляющий правами доступа к папкам
    /// </summary>
    public static class AccessSetter
    {
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
        public static bool RemoveUserAccess(DirectoryInfo root, string userNameRaw)
        {
            bool result = true;
            string user = GetUsernamesFromString(userNameRaw).FirstOrDefault();
            if (root.Exists)
            {
                string[] subDirs = root.GetSubDirsRecursively().Reverse().ToArray();
                for (int i = 0; i < subDirs.Length; i++)
                {
                    result &= SetDirAccessForUser(new UserDirAccessDto(subDirs[i], user, UserAccessRule.Disable));
                }
                return result;
            }
            else
            {
                DirectoryNotExistMessage(root.FullName);
                return false;
            }
        }

        /// <summary>
        /// Возвращает массив валидных имен пользователей с доменными префиксами из ячейки.
        /// </summary>
        /// <param name="s">Строковое значение ячейки</param>
        /// <returns>Массив имен пользователей с доменным префиксом</returns>
        public static string[] GetUsernamesFromString(string s)
        {
            Regex regex = new Regex(@"[a-z]+[.][a-z]{2}$");
            var usersRaw = s.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            List<string> users = new List<string>();
            foreach (var user in usersRaw)
            {
                var u = user.StartsWith("\n") ? user.Remove(0, 1) : user;
                if (regex.IsMatch(u))
                {
                    users.Add(@"PGS.ru\" + u);
                }
            }
            return users.ToArray();
        }

        /// <summary>
        /// Назначает права доступа к папке для пользователя.
        /// Если папка не существует, она будет предварительно создана.
        /// </summary>
        /// <param name="userDirAccess">Dto с правами доступа к папке для пользователя</param>
        /// <returns>True, если назначение прав прошло успешно, иначе false</returns>
        public static bool SetDirAccessForUser(UserDirAccessDto userDirAccess)
        {
            try
            {
                UserAccessRule rule = userDirAccess.Access;
                if (!Directory.Exists(userDirAccess.DirPath))
                {
                    Directory.CreateDirectory(userDirAccess.DirPath);
                    Console.WriteLine(new string('=', 140));
                    Console.WriteLine($"=== Create --> {userDirAccess.DirPath,-121} ===", -140);
                    Console.WriteLine(new string('=', 140));
                }
                Console.WriteLine($"\t{userDirAccess.Access,-10}\t{userDirAccess.UserName,-20}\t{userDirAccess.DirPath}");

                switch (rule)
                {
                    case UserAccessRule.Read:
                        return SetReadRules(userDirAccess.DirPath, userDirAccess.UserName, false);
                    case UserAccessRule.ReadDeep:
                        return SetReadRules(userDirAccess.DirPath, userDirAccess.UserName, true);
                    case UserAccessRule.Write:
                        return SetWriteRules(userDirAccess.DirPath, userDirAccess.UserName);
                    case UserAccessRule.Disable:
                        return RemoveAccess(userDirAccess.DirPath, userDirAccess.UserName);
                    default:
                        return false;
                }
            }
            catch (IdentityNotMappedException)
            {
                Console.WriteLine($"Ошибка: имя пользователя {userDirAccess.UserName} некорректно !");
                return false;
            }
        }


        /// <summary>
        /// Удаляет доступ пользователя к заданной директории и всем ее подпапкам
        /// </summary>
        /// <param name="DirPath">Путь к корневой папке</param>
        /// <param name="UserName">Имя пользователя с доменным префиксом</param>
        /// <returns>True, если команда выполнена успешно, иначе false</returns>
        private static bool RemoveAccess(string DirPath, string @UserName)
        {
            bool result = true;
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, true);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, true);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, true);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, true);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, true);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, true);
            return result;
        }

        /// <summary>
        /// Назначает доступ на редактирование для заданной папки, ее подпапок и файлов
        /// </summary>
        /// <param name="DirPath">Полный путь к заданной папке</param>
        /// <param name="UserName">Имя пользователя с доменным префиксом</param>
        /// <returns>True, если команда выполнена успешно, иначе false</returns>
        private static bool SetWriteRules(string DirPath, string @UserName)
        {
            bool result = true;
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, true);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, true);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, true);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, true);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, true);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, true);
            return result;
        }


        /// <summary>
        /// Назначает доступ на чтение
        /// </summary>
        /// <param name="DirPath">Полный путь к заданной папке</param>
        /// <param name="UserName">Имя пользователя с доменным префиксом</param>
        /// <param name="isDeepAccess">
        /// Если True => для заданной папки ее подпапок и файлов;
        /// Если False => только для заданной папки
        /// </param>
        /// <returns>True, если назначение прав выполнено успешно, иначе false</returns>
        private static bool SetReadRules(string DirPath, string @UserName, bool isDeepAccess)
        {
            bool result = true;
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.FullControl, isDeepAccess);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Modify, isDeepAccess);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ReadAndExecute, isDeepAccess);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.ListDirectory, isDeepAccess);
            result &= AddDirectorySecurity(DirPath, @UserName, FileSystemRights.Read, isDeepAccess);
            result &= RemoveDirectorySecurity(DirPath, @UserName, FileSystemRights.Write, isDeepAccess);
            return result;
        }

        /// <summary>
        /// Сообщение об ошибке, если директории по заданному пути не существует
        /// </summary>
        /// <param name="DirPath">Полный путь к заданной директории</param>
        private static void DirectoryNotExistMessage(string DirPath)
        {
            Console.WriteLine($"Ошибка: папки {DirPath} не существует !");
        }

        /// <summary>
        /// Назначает разрешение заданного действия к заданной директории
        /// </summary>
        /// <param name="DirPath">Полный путь к директории</param>
        /// <param name="UserName">Имя пользователя с доменным префиксом</param>
        /// <param name="Rights">Действие, которое нужно разрешить пользователю</param>
        /// <param name="isDeepAccess">
        /// Если True => для заданной папки ее подпапок и файлов;
        /// Если False => только для заданной папки
        /// </param>
        /// <returns>True, если назначение прав выполнено успешно, иначе false</returns>
        private static bool AddDirectorySecurity(
            string DirPath,
            string UserName,
            FileSystemRights Rights,
            bool isDeepAccess)
        {
            DirectoryInfo dInfo = new DirectoryInfo(DirPath);
            if (!dInfo.Exists)
            {
                DirectoryNotExistMessage(DirPath);
                return false;
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
            return true;
        }

        /// <summary>
        /// Запрещает заданное действие с заданной папкой
        /// </summary>
        /// <param name="DirPath">Полный путь к директории</param>
        /// <param name="UserName">Имя пользователя с доменным префиксом</param>
        /// <param name="Rights">Действие, которое нужно запретить для пользователя</param>
        /// <param name="isDeepAccess">
        /// Если True => для заданной папки ее подпапок и файлов;
        /// Если False => только для заданной папки
        /// </param>
        /// <returns>True, если назначение прав выполнено успешно, иначе false</returns>
        private static bool RemoveDirectorySecurity(
            string DirPath,
            string UserName,
            FileSystemRights Rights,
            bool isDeepAccess)
        {
            DirectoryInfo dInfo = new DirectoryInfo(DirPath);
            if (!dInfo.Exists)
            {
                DirectoryNotExistMessage(DirPath);
                return false;
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
            return true;
        }
    }
}
