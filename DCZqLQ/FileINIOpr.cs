using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace KY.Fi.DCZqLQ
{
    class FileINIOpr
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);
        
        /// <summary>
        /// 写入值
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="val"></param>
        /// <param name="filepath"></param>
        public void SetIniKeyValue(string section, string key, string val, string filepath)
        {
            WritePrivateProfileString(section, key, val, filepath);
        }

        /// <summary>
        /// 读取值
        /// </summary>
        /// <param name="Section"></param>
        /// <param name="key"></param>
        /// <param name="strFilePath"></param>
        /// <returns></returns>
        public string GetIniKeyValue(string Section, string key, string strFilePath)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

    }
}
