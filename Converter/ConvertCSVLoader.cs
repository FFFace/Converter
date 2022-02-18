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
                CreateCSUtil.CSTabWriteLine(sw, $"public Dictionary<int, {fileName}> GetDic()");
                CreateCSUtil.CSTabWriteLine(sw, "{");
                CreateCSUtil.AddTabCount();
                CreateCSUtil.CSTabWriteLine(sw, $"string path = Application.dataPath + \"/data/{fileName}.csv\";");
                CreateCSUtil.CSTabWriteLine(sw, $"Dictionary<int, {fileName}> Dic = new Dictionary<int, {fileName}>();");
                CreateCSUtil.CSTabWriteLine(sw, "");
                CreateCSUtil.CSTabWriteLine(sw, "StreamReader sr = new StreamReader(path);");
                CreateCSUtil.CSTabWriteLine(sw, $"string[] lines = sr.ReadToEnd().Split(\'\\n\');");
                CreateCSUtil.CSTabWriteLine(sw, "");
                CreateCSUtil.CSTabWriteLine(sw, "for(int i=4; i<lines.Length; i++)");
                CreateCSUtil.CSTabWriteLine(sw, "{");
                CreateCSUtil.AddTabCount();
                CreateCSUtil.CSTabWriteLine(sw, $"{fileName} item = new {fileName}();");
                CreateCSUtil.CSTabWriteLine(sw, "string[] cells = lines[i].Split(\',\');");

                Range range = sheet.UsedRange;
                string name, type;
                for (int column = 1; column <= range.Columns.Count; column++)
                {
                    name = Convert.ToString((range.Cells[1, column] as Range).Value2);
                    type = Convert.ToString((range.Cells[2, column] as Range).Value2);

                    switch (type)
                    {
                        case "int":
                            CreateCSUtil.CSTabWriteLine(sw, $"item.{name} = Convert.ToInt32(cells[{column}]);");
                            break;

                        case "float":
                            CreateCSUtil.CSTabWriteLine(sw, $"item.{name} = Convert.ToSingle(cells[{column}]);");
                            break;

                        case "string":
                            CreateCSUtil.CSTabWriteLine(sw, $"item.{name} = Convert.ToString(cells[{column}]);");
                            break;

                        case "bool":
                            CreateCSUtil.CSTabWriteLine(sw, $"item.{name} = Convert.ToBoolean(cells[{column}];");
                            break;

                        default:
                            if (type.Contains("["))
                            {
                                type = type.Remove(type.IndexOf('['));
                                CreateCSUtil.CSTabWriteLine(sw, $"string[] array = cells[{column}].Split(';');");
                                CreateCSUtil.CSTabWriteLine(sw, $"item.{name} = new {type}[array.Length];");
                                CreateCSUtil.CSTabWriteLine(sw, $"for(int j = 0; i < array.Length; i++)");
                                CreateCSUtil.CSTabWriteLine(sw, "{");
                                CreateCSUtil.AddTabCount();
                                switch (type)
                                {
                                    case "int":
                                        CreateCSUtil.CSTabWriteLine(sw, $"item.{name}[i] = Convert.ToInt32(array[i]);");
                                        break;

                                    case "float":
                                        CreateCSUtil.CSTabWriteLine(sw, $"item.{name}[i] = Convert.ToSingle(array[i]);");
                                        break;

                                    case "string":
                                        CreateCSUtil.CSTabWriteLine(sw, $"item.{name}[i] = Convert.ToInt32(array[i]);");
                                        break;

                                    case "bool":
                                        CreateCSUtil.CSTabWriteLine(sw, $"item.{name}[i] = Convert.ToBoolean(array[i]);");
                                        break;

                                    default:

                                        break;
                                }
                                CreateCSUtil.RemoveTabCount();
                                CreateCSUtil.CSTabWriteLine(sw, "}");
                            }
                            break;
                    }
                }
                CreateCSUtil.CSTabWriteLine(sw, "Dic.Add(item.index, item);");
                CreateCSUtil.RemoveTabCount();
                CreateCSUtil.CSTabWriteLine(sw, "}");
                CreateCSUtil.CSTabWriteLine(sw, "return Dic;");
                CreateCSUtil.CSTabClose(sw);
            }
            file.Close();
        }
    }
}
