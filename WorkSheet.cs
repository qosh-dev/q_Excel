using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel
{
    public class WorkSheet
    {
        public int fontSize = 12;
        private ExcelPackage package;
        public ExcelWorksheet worksheet;
        public string worksheetName;
        private int row = 1;
        private int colomn = 1;

        public WorkSheet(ExcelPackage package, string worksheetName)
        {
            this.package = package;
            this.worksheetName = worksheetName;
            this.worksheet = package.Workbook.Worksheets.Add(worksheetName);
        }

        public ExcelRange Params() => worksheet.Cells[worksheet.Dimension.Address];

        public ExcelRange this[int row, int column] => this.worksheet.Cells[row, column];

        public ExcelRange this[int row, int column, int row2, int column2] => worksheet.Cells[row, column, row2, column2];

        public WorkSheet AddHeaders<T>(T[] headersList, int colomn = 1, int row = 0, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Left, bool isBold = true, int fontSize = 12)
        {
            if (row == 0)
            {
                row = this.row;
                this.row++;
            }
            for (int i = 0; i < headersList.Count(); i++)
            {
                worksheet.Cells[row, i + colomn].Value = headersList[i];
                worksheet.Cells[row, i + colomn].Style.Font.Size = fontSize;
                worksheet.Cells[row, i + colomn].Style.Font.Bold = isBold;
                worksheet.Cells[row, i + colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                worksheet.Cells[row, i + colomn].Style.HorizontalAlignment = horizontalAlignment;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

        public WorkSheet AddMergedHeaders(string[] valueIndexArr, int step, int row = 0, int colomn = 0, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center, bool isBold = true, int fontSize = 12)
        {
            var arr = Excel.Merge.buildMerge(valueIndexArr, step);
            if (row == 0)
            {
                row = this.row;
                this.row++;
            }
            colomn -= 1;
            for (int i = 0; i < valueIndexArr.Count(); i++)
            {
                var el = arr[i];
                worksheet.Cells[row, el.indStart + colomn, row, el.indStop + colomn].Merge = true;
                worksheet.Cells[row, el.indStart + colomn, row, el.indStop + colomn].Value = el.val;
                worksheet.Cells[row, el.indStart + colomn, row, el.indStop + colomn].Style.Font.Size = fontSize;
                worksheet.Cells[row, el.indStart + colomn, row, el.indStop + colomn].Style.Font.Bold = isBold;
                worksheet.Cells[row, el.indStart + colomn, row, el.indStop + colomn].Style.HorizontalAlignment = horizontalAlignment;
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

        public WorkSheet AddLoop(List<string> list, int colomn = 1, int row = 0)
        {

            for (int i = 0; i < list.Count(); i++)
            {
                if (row == 0)
                {
                    worksheet.Cells[this.row, colomn].Style.Font.Size = fontSize;
                    worksheet.Cells[this.row, colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    worksheet.Cells[this.row, colomn].Value = list[i];
                    this.row++;
                }
                else
                {
                    worksheet.Cells[row, colomn].Style.Font.Size = fontSize;
                    worksheet.Cells[row, colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    worksheet.Cells[row, colomn].Value = list[i];
                    row++;
                }
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }
        public WorkSheet AddLoop(List<int> list, int colomn = 1, int row = 0)
        {

            for (int i = 0; i < list.Count(); i++)
            {
                if (row == 0)
                {
                    worksheet.Cells[this.row, colomn].Style.Font.Size = fontSize;
                    worksheet.Cells[this.row, colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    worksheet.Cells[this.row, colomn].Value = list[i];
                    this.row++;
                }
                else
                {
                    worksheet.Cells[row, colomn].Style.Font.Size = fontSize;
                    worksheet.Cells[row, colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    worksheet.Cells[row, colomn].Value = list[i];
                    row++;
                }
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

        public WorkSheet AddLoop<T>(List<T> list, int colomn = 1, int row = 0)
        {

            string[] propList = typeof(T).GetProperties().Select(e => e.Name).ToArray<string>();
            var tempList = Methods.toJson(list);

            for (int i = 0; i < tempList.Count(); i++)
            {
                if (row == 0)
                {
                    for (int o = 0; o < propList.Count(); o++)
                    {
                        worksheet.Cells[this.row, o + colomn].Style.Font.Size = fontSize;
                        worksheet.Cells[this.row, o + colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[this.row, o + colomn].Value = tempList[i][propList[o]];
                    }
                    this.row++;
                }
                else
                {
                    for (int o = 0; o < propList.Count(); o++)
                    {
                        worksheet.Cells[row, o + colomn].Style.Font.Size = fontSize;
                        worksheet.Cells[row, o + colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[row, o + colomn].Value = tempList[i][propList[o]];
                    }
                    row++;
                }
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

        public WorkSheet AddLoopVertical<T>(List<T> list, int colomn = 1, int row = 0)
        {
            string[] propList = typeof(T).GetProperties().Select(e => e.Name).ToArray<string>();
            int range = list[0].GetType().GetProperties().Select(e => e.Name).Count();
            var tempList = Methods.toJson(list);

            for (int i = 0; i < list.Count(); i++)
            {
                if (colomn == 0)
                {
                    for (int o = 0; o < propList.Count(); o++)
                    {
                        worksheet.Cells[this.row, o + colomn].Style.Font.Size = fontSize;
                        worksheet.Cells[this.row, o + colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[this.row, o + colomn].Value = tempList[i][propList[o]];
                    }
                    this.colomn += range;
                }
                else
                {
                    for (int o = 0; o < propList.Count(); o++)
                    {
                        worksheet.Cells[row, o + colomn].Style.Font.Size = fontSize;
                        worksheet.Cells[row, o + colomn].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[row, o + colomn].Value = tempList[i][propList[o]];
                    }
                    colomn += range;
                }
            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

    }

}
