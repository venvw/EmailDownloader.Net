using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace EmailDownloader.Models
{
    internal static class Debug
    {
        public static string OutputPath = $"{Environment.CurrentDirectory}//output_log.txt";

        private static List<string> lines;

        public static void Log(object a)
        {
            if (lines == null)
                lines = new List<string>();
            
            lines.Add(a.ToString());
            File.WriteAllLines(OutputPath, lines);
        }
    }
}
