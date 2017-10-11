using System;
using System.Web;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Text;
using System.IO;

namespace TestGetFiles
{
    /// <summary>
    /// bsGetFiles 的摘要描述。
    /// </summary>
    public class bsGetFiles
    {
        public bsGetFiles()
        {
        }
        private static string result = "";
        /// <summary>
        /// 得某文件夹下所有的文件
        /// </summary>
        /// <param name="directory">文件夹名称</param>
        /// <param name="pattern">搜寻指类型</param>
        /// <returns></returns>
        public static string GetFiles(DirectoryInfo directory, string pattern)
        {
            if (directory.Exists || pattern.Trim() != string.Empty)
            {

                foreach (FileInfo info in directory.GetFiles(pattern))
                {
                    result = result + info.FullName.ToString() + ";";
                    //result = result + info.Name.ToString() + ";";
                }

                foreach (DirectoryInfo info in directory.GetDirectories())
                {
                    GetFiles(info, pattern);
                }
            }
            string returnString = result;
            return returnString;

        }

    }
}