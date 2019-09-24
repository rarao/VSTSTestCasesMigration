using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSTSTestCasesMigration
{
    public static class Helpers
    {
        public static string ReplaceReservedChars(string str)
        {
            string[] strArr = new string[] { "<", ">", ":", "\"", "/", "\\", "|", "?", "*","&" };

            for (int i = 0; i < strArr.Length; i++)
            {
                str = str.Replace(strArr[i], "");
            }

            string returnstr = new string(str.Where(c => c < 128).ToArray());

            return returnstr;
        }
    }
}
