using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Converter
{
    class CreateCSUtil
    {
        private static int tabCount = 0;
        public static void CSHeaderWirte(StreamWriter sw, string fileName, string nameSpace = "")
        {
            sw.WriteLine("using System;");
            sw.WriteLine("");
            if (nameSpace != "")
            {
                sw.WriteLine($"namespace {nameSpace}");
                sw.WriteLine("{");
                tabCount += 1;
            }

            CSTabWriteLine(sw, $"public class {fileName}");
            CSTabWriteLine(sw, "{");
            tabCount += 1;
        }

        public static void CSHeaderWirteForUnity(StreamWriter sw, string fileName, string nameSpace = "")
        {
            sw.WriteLine("using System;");
            sw.WriteLine("");
            if (nameSpace != "")
            {
                sw.WriteLine($"namespace {nameSpace}");
                sw.WriteLine("{");
                tabCount += 1;
            }

            CSTabWriteLine(sw, $"public class {fileName} : ScriptableObject");
            CSTabWriteLine(sw, "{");
            tabCount += 1;
        }

        public static void CSTabWriteLine(StreamWriter sw, string str)
        {
            for (int i = 0; i < tabCount; i++)
            {
                sw.Write("\t");
            }

            sw.WriteLine(str);
        }

        public static void CSTabWrite(StreamWriter sw, string str)
        {
            for (int i = 0; i < tabCount; i++)
            {
                sw.Write("\t");
            }

            sw.Write(str);
        }

        public static void CSTabClose(StreamWriter sw)
        {
            for (int i = tabCount; i > 0; i--)
            {
                for (int j = 0; j < i - 1; j++)
                {
                    sw.Write("\t");
                }

                sw.WriteLine("}");
            }

            tabCount = 0;
        }

        public static void AddTabCount()
        {
            tabCount += 1;
        }

        public static void RemoveTabCount()
        {
            tabCount -= 1;
        }
    }
}
