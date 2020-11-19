using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel
{
    public class Create 
    {
        public byte[] ExcelResult(Action<ExcelPackage> result)
        {
            var package = new ExcelPackage();
            result(package); 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            return package.GetAsByteArray();
        }

        public void Build(Action<ExcelPackage> result, string path)
        {
            //Build new Excel file
            var package = new ExcelPackage();
            result(package);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            package.SaveAs( path switch {
                "." => new FileInfo($@"{Directory.GetCurrentDirectory()}\\worksheet.xlsx".Replace("\'", "\\'")),
                _ => new FileInfo($@"{path}.xlsx".Replace("\'", "\\'"))
            });
        }


        
    }

    public static class ExcelResult {
        public static FileContentResult Excel(this ControllerBase obj, Action<ExcelPackage> result,string fileName = "file"){
            var package = new ExcelPackage();
            result(package);
            return obj.File(
                fileContents: package.GetAsByteArray(),
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: fileName + ".xlsx");
        }

        public static WorkSheet addWorkSheet(this ExcelPackage obj, string worksheetName){
            return new WorkSheet(obj,worksheetName);
        }
    }
}
