using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Converter
{
    class ConvertCS
    {
        private string filePath;
        private string fileName;
        private Worksheet sheet;
        private string nameSpace = "";

        public ConvertCS() { }

        public ConvertCS(Worksheet _sheet, string _filePath, string _fileName, string _nameSpace = "")
        {
            filePath = _filePath;
            fileName = _fileName;
            sheet = _sheet;
            nameSpace = _nameSpace;
        }

        public void SetFilePath(string _filePath)
        {
            filePath = _filePath;
        }

        public void SetFileName(string _fileName)
        {
            fileName = _fileName;
        }

        public void CreateCS()
        {
            if (filePath == null)
            {
                Console.WriteLine("File Path Is Null");
                return;
            }

            if (!Directory.Exists(filePath))
            {
                Console.WriteLine("File Path Is Wrong");
                return;
            }

            FileStream file = File.Create(Path.Combine(filePath, fileName + ".cs"));
            using (StreamWriter sw = new StreamWriter(file, Encoding.UTF8))
            {
                CreateCSUtil.CSHeaderWirte(sw, fileName, nameSpace);
                Range range = sheet.UsedRange;
                for (int column = 1; column <= range.Columns.Count; column++)
                {
                    string name = Convert.ToString((range.Cells[1, column] as Range).Value2);
                    string type = Convert.ToString((range.Cells[2, column] as Range).Value2);
                    string explain = Convert.ToString((range.Cells[3, column] as Range).Value2);

                    CreateCSUtil.CSTabWrite(sw, $"public {type} {name}; //{explain}");
                }

                CreateCSUtil.CSTabClose(sw);
                
            }
            file.Close();
        }
    }
}
