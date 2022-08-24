using DirsAccessFromExcel.Logic;

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
