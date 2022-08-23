using DirsAccessFromExcele.Logic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DirsAccessFromExcele.DTO
{
    public class UserDirAccessDto
    {
        public string @DirPath { get; private set; }

        public string @UserName { get; private set; }

        public UserAccessRule Access { get; private set; }

        public bool IsDeepAccess { get; private set; }

        public UserDirAccessDto(string DirPath, string UserName, UserAccessRule Access, bool isDeepAccess = false)
        {
            this.DirPath = DirPath;
            this.UserName = UserName;
            this.Access = Access;
            if (Access != UserAccessRule.Disable)
            {
                IsDeepAccess = isDeepAccess;
            }
        }
    }
}
