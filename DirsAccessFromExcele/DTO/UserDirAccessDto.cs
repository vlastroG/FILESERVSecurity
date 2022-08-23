using DirsAccessFromExcel.Logic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcel.DTO
{
    public class UserDirAccessDto
    {
        public string @DirPath { get; private set; }

        public string @UserName { get; private set; }

        public UserAccessRule Access { get; private set; }

        public UserDirAccessDto(string DirPath, string UserName, UserAccessRule Access)
        {
            this.DirPath = DirPath;
            this.UserName = UserName;
            this.Access = Access;
        }
    }
}
