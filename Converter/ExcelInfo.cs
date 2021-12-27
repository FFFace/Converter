using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Converter
{
    class ExcelInfo
    {
        private Application excel = null;
        private Workbook workBook = null;
        private Worksheet workSheet = null;

        private string filePath = null;
        private int count = 0;

        public ExcelInfo() { }
        public ExcelInfo(string _filePath)
        {
            filePath = _filePath;
        }

        public bool ReadFileFromFilePath()
        {
            if (filePath == null)
            {
                Console.WriteLine("File Path Is Null. Please Set File Path. (Method : SetFilePath(string))");
                return false;
            }

            excel = new Application();
            workBook = excel.Workbooks.Open(filePath);
            bool isOpen = false;

            if (workBook == null)
            {
                isOpen = false;
                Console.WriteLine("Can Not Found File. Please Check File Path");
                return isOpen;
            }

            isOpen = true;
            Console.WriteLine("Found File.");

            if (workBook.Worksheets.Count > 0)
                workSheet = workBook.Worksheets[1];

            else
                Console.WriteLine($"Sheet Count Is {workBook.Worksheets.Count}");
            return isOpen;
        }

        public void SetFilePath(string _filePath)
        {
            filePath = _filePath;
        }

        public Worksheet GetWorkSheet()
        {
            if (workSheet == null)
            {
                Console.WriteLine("WorkSheet Is Null");
                return null;
            }

            return workSheet;
        }

        public void ExcelClose()
        {
            Marshal.ReleaseComObject(workSheet);
            workBook.Close();
            excel.Quit();            
        }
    }
}
