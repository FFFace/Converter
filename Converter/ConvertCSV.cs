using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace Converter
{
    class ConvertCSV
    {
        private Worksheet sheet;
        private string csvFilePath;
        private string fileName;

        public ConvertCSV() { }

        public ConvertCSV(Worksheet _sheet, string _csvFilePath, string _fileName)
        {
            sheet = _sheet;
            csvFilePath = _csvFilePath;
            fileName = _fileName;
        }

        public void SetSheet(Worksheet _sheet)
        {
            sheet = _sheet;
        }

        public void SetCSVFilePath(string _csvFilePath)
        {
            csvFilePath = _csvFilePath;
        }

        public void SetFileName(string _fileName)
        {
            fileName = _fileName;
        }

        public void CreateCSV()
        {
            if (sheet == null)
            {
                Console.WriteLine("Sheet Is Null");
                return;
            }

            if (!Directory.Exists(csvFilePath))
                Directory.CreateDirectory(csvFilePath);

            FileStream file = File.Create(Path.Combine(csvFilePath, fileName + ".csv"));
            using (StreamWriter sw = new StreamWriter(file, Encoding.UTF8))
            {

                Range range = sheet.UsedRange;
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int column = 1; column <= range.Columns.Count; column++)
                    {
                        Range cell = range.Cells[row, column];
                        string str = Convert.ToString(cell.Value2) + ",";
                        sw.Write(str);
                    }
                    sw.Write("\n");
                }
            }
            file.Close();
        }
    }
}
