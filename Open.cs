using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Newtonsoft.Json.JsonConvert;

namespace Excel
{
    public class Open
    {
        private string fileName;
        public bool conteinHeaders = false;
        public string ExceptionMessege = "Incorrect signature of Excel data";

        public Open(string fileName, bool isCurrentDirectory = true) => this.fileName = isCurrentDirectory ? $"{Directory.GetCurrentDirectory()}\\" + fileName : fileName;

        public string this[int row, int column]
        {
            get
            {
                byte[] bin = File.ReadAllBytes(fileName);
                using (MemoryStream stream = new MemoryStream(bin))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(stream))
                    {
                        ExcelWorksheet obj = excelPackage.Workbook.Worksheets[0];
                        return obj.Cells[row, column].Value.ToString();
                    }
                }
            }
        }
        public dynamic ParseRow<T>()
        {
            byte[] bin = File.ReadAllBytes(fileName);
            using MemoryStream stream = new MemoryStream(bin);
            using ExcelPackage excelPackage = new ExcelPackage(stream);
            ExcelWorksheet obj = excelPackage.Workbook.Worksheets[0];
            return this.Parse<T>(obj, 2);
        }

        public List<T> ParseTo<T>(int row = 0)
        {
            byte[] bin = File.ReadAllBytes(fileName);
            using MemoryStream stream = new MemoryStream(bin);
            using ExcelPackage excelPackage = new ExcelPackage(stream);
            ExcelWorksheet obj = excelPackage.Workbook.Worksheets[0];
            var list = new List<T>();
            var i = row;
            while (true)
            {
                try
                {
                    i++;
                    list.Add(this.Parse<T>(obj, i));
                }
                catch (System.Exception)
                {
                    break;
                }

            }
            return list;

        }

        private dynamic Parse<T>(ExcelWorksheet reader, int row)
        {
            string typeString = @"{";
            var props = typeof(T).GetProperties();
            for (int i = 1; i < props.Count(); i++)
            {
                typeString += $" '{props[i - 1].Name}': '{reader.Cells[row, i].Value}' ,";
            }
            typeString += "}";
            return DeserializeObject<T>(typeString);
        }


    }
}
