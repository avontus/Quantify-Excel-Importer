using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Configuration;

namespace Avontus.Quantify.ExcelImporter
{
    public static class ExensionMethods
    {
        
        public static string Left(this string s, int numOfChars)
        {
            if (numOfChars >= s.Length)
                return s;
            else
                return s.Substring(0, numOfChars);
        }

        public static string Right(this string s, int numOfChars)
        {
            if (numOfChars >= s.Length)
                return s;
            else
                return s.Substring(s.Length - numOfChars, numOfChars);
        }

        public static bool IsNotNullOrEmpty(this string str)
        {
            return !string.IsNullOrEmpty(str);
        }

    }
}
