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
    class ConvertCSVLoader
    {
        private Worksheet sheet;
        private string filePath;
        private string fileName;
        private string nameSpace = "";

        public ConvertCSVLoader() { }
        public ConvertCSVLoader(Worksheet _sheet, string _filePath, string _fileName, string _nameSpace = "")
        {
            sheet = _sheet;
            filePath = _filePath;
            fileName = _fileName;
            nameSpace = _nameSpace;
        }

        public void SetWorksheet(Worksheet _sheet)
        {
            sheet = _sheet;
        }

        public void SetFilepath(string _filePath)
        {
            filePath = _filePath;
        }

        public void SetFileName(string _fileName)
        {
            fileName = _fileName;
        }

        public void SetNameSpaceE(string _nameSpace)
        {
            nameSpace = _nameSpace;
        }

        public void CreateCSVLoader()
        {
            if (filePath == null)
            {
                Console.WriteLine("File Path Is Null");
                return;
            }

            if (!Directory.Exists(filePath))
            {
                Console.WriteLine("File Path is Null");
                return;
            }

            FileStream file = File.Create(Path.Combine(filePath, fileName + "Loader.cs"));
            using (StreamWriter sw = new StreamWriter(file, Encoding.UTF8))
            {
                sw.WriteLine("using UnityEngine;");
                sw.WriteLine("using System.Collections.Generic;");
                sw.WriteLine("using System.IO;");
                CreateCSUtil.CSHeaderWirte(sw, fileName + "Loader", nameSpace);
                CreateCSUtil.CSTabWrite(sw, $"public Dictionary<int, {fileName}> GetDic()");
                CreateCSUtil.CSTabWrite(sw, "{");
                CreateCSUtil.AddTabCount();
                CreateCSUtil.CSTabWrite(sw, $"string path = Application.dataPath + \"/data/{fileName}.csv\";");
                CreateCSUtil.CSTabWrite(sw, $"Dictionary<int, {fileName}> Dic = new Dictionary<int, {fileName}>();");
                CreateCSUtil.CSTabWrite(sw, "");
                CreateCSUtil.CSTabWrite(sw, "StreamReader sr = new StreamReader(path);");
                CreateCSUtil.CSTabWrite(sw, $"string[] lines = sr.ReadToEnd().Split(\'\\n\');");
                CreateCSUtil.CSTabWrite(sw, $"{fileName} item = new {fileName}();");
                CreateCSUtil.CSTabWrite(sw, "");
                CreateCSUtil.CSTabWrite(sw, "for(int i=4; i<lines.Length; i++)");
                CreateCSUtil.CSTabWrite(sw, "{");
                CreateCSUtil.AddTabCount();
                CreateCSUtil.CSTabWrite(sw, "string[] cells = lines[i].Split(\',\');");

                Range range = sheet.UsedRange;
                string name, type;
                for (int column = 1; column <= range.Columns.Count; column++)
                {
                    name = Convert.ToString((range.Cells[1, column] as Range).Value2);
                    type = Convert.ToString((range.Cells[2, column] as Range).Value2);

                    switch (type)
                    {
                        case "int":
                            CreateCSUtil.CSTabWrite(sw, $"item.{name} = Convert.ToInt32(cells[{column}]);");
                            break;

                        case "float":
                            CreateCSUtil.CSTabWrite(sw, $"item.{name} = Convert.ToSingle(cells[{column}]);");
                            break;

                        case "string":
                            CreateCSUtil.CSTabWrite(sw, $"item.{name} = Convert.ToString(cells[{column}]);");
                            break;

                        case "bool":
                            CreateCSUtil.CSTabWrite(sw, $"item.{name} = {column};");
                            break;

                        default:

                            break;
                    }
                }
                CreateCSUtil.CSTabWrite(sw, "Dic.Add(item.index, item);");
                CreateCSUtil.RemoveTabCount();
                CreateCSUtil.CSTabWrite(sw, "}");
                CreateCSUtil.CSTabWrite(sw, "return Dic;");
                CreateCSUtil.CSTabClose(sw);
            }
            file.Close();
        }
    }
}
