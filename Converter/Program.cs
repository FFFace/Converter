using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Converter
{
    class Program
    {
        private const int argsCount = 2;

        /// <summary>
        /// args[0] = ProjectName
        /// args[1] = ExcelName
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (args.Length != argsCount)
            {
                Console.WriteLine("CommandLine Count Fail. Check  CommandLine.");
                ProgramEnd();
                return;
            }

            string projectName = args[0];
            string fileName = args[1];

            string filePath = Directory.GetCurrentDirectory();
            for (int i = 0; i < 5; i++)
               filePath = Path.GetDirectoryName(filePath);

            string createCSVFilePath = Path.Combine(filePath, projectName, "Assets", "Data");
            string createCSFilePath = Path.Combine(filePath, projectName, "Assets", "Scripts", "Data");
            string createCSVLoaderFilePath = Path.Combine(filePath, projectName, "Assets", "Scripts", "DataLoader");

            if (!Directory.Exists(createCSFilePath))
                Directory.CreateDirectory(createCSFilePath);

            if (!Directory.Exists(createCSVLoaderFilePath))
                Directory.CreateDirectory(createCSVLoaderFilePath);

            ExcelInfo info = new ExcelInfo();
            info.SetFilePath(Path.Combine(filePath, "Excel", fileName + ".xlsx"));
            if (!info.ReadFileFromFilePath())
            {
                info.ExcelClose();
                ProgramEnd();
                return;
            }

            ConvertCSV convertCSV = new ConvertCSV(info.GetWorkSheet(), createCSVFilePath, fileName);
            convertCSV.CreateCSV();

            ConvertCS convertCS = new ConvertCS(info.GetWorkSheet(), createCSFilePath, fileName, "Test");
            convertCS.CreateCS();

            ConvertCSVLoader convertCSVLoader = new ConvertCSVLoader(info.GetWorkSheet(), createCSVLoaderFilePath, fileName, "Test");
            convertCSVLoader.CreateCSVLoader();

            info.ExcelClose();

            ProgramEnd();
        }

        private static void ProgramEnd()
        {
            Console.WriteLine("Please Input Enter...");
            Console.ReadLine();
        }
    }
}
