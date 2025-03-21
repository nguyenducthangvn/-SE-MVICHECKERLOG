using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class iniHelper
    {
        private string _filePath;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        public iniHelper(string filePath)
        {
            _filePath = filePath;
        }

        public string ReadValue(string section, string key, string defaultValue = "")
        {
            const int bufferSize = 255;
            var buffer = new StringBuilder(bufferSize);
            GetPrivateProfileString(section, key, defaultValue, buffer, bufferSize, _filePath);
            return buffer.ToString();
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        public bool WriteValue(string section, string key, string value)
        {
            return WritePrivateProfileString(section, key, value, _filePath);
        }
    }
}
