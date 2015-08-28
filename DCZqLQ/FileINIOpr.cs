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

        public void SetIniKeyValue(string section, string key, string val, string filepath)
        {
            WritePrivateProfileString(section, key, val, filepath);
        }

        public string GetIniKeyValue(string src, string key, string strFilePath)
        {            
            return ContentValue(src, key, strFilePath);
        }

        private string ContentValue(string Section, string key, string strFilePath)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }
    }
}
