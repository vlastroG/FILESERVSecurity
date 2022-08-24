using System;
using System.Collections.Generic;
using System.IO;

namespace DirsAccessFromExcel.DTO
{
    public static class DirectoryInfoExtension
    {
        /// <summary>
        /// Возвращает все подкаталоги из корневой папки в порядке от внешней к внутренней,
        /// порядок подпапок верхнего уровня не соблюдается.
        /// </summary>
        /// <param name="root">Корневой каталог</param>
        /// <returns>Массив полных путей к подкаталогам</returns>
        public static string[] GetSubDirsRecursively(this DirectoryInfo root)
        {
            List<string> subDirs = new List<string>();
            if (root.Exists)
            {
                FillSubdirsListRecurs(root, ref subDirs);
                return subDirs.ToArray();
            }
            else
            {
                return new string[0];
            }
        }

        private static void FillSubdirsListRecurs(DirectoryInfo root, ref List<string> pathList)
        {
            DirectoryInfo[] subDirs = null;
            try
            {
                subDirs = root.GetDirectories();
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
            foreach (DirectoryInfo dirInfo in subDirs)
            {
                pathList.Add(dirInfo.FullName);
                FillSubdirsListRecurs(dirInfo, ref pathList);
            }
        }
    }
}
