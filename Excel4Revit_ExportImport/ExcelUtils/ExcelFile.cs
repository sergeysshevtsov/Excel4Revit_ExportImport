using Excel4Revit_ExportImport.RevitExtensions;
using Excel4Revit_ExportImport.Views.ViewExport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace Excel4Revit_ExportImport.ExcelUtils;
public static class ExcelFile
{
    public static void Export(List<Element> elements, List<ParameterData> parametersData)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Excel4Revit_ExportImport");
        string filePath = Path.Combine(Path.GetTempPath(), "ExcelExport.xlsx");

        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("RevitExport");

        // Add headers
        worksheet.Cells[1, 1].Value = "Name";
        worksheet.Cells[1, 2].Value = "Family type";
        worksheet.Cells[1, 3].Value = "Category";
        worksheet.Cells[1, 4].Value = "Location X";
        worksheet.Cells[1, 5].Value = "Location Y";
        worksheet.Cells[1, 6].Value = "Location Z";
        worksheet.Cells[1, 7].Value = "Rotation";

        for (int i = 0; i < parametersData.Count; i++)
            worksheet.Cells[1, 8 + i].Value = parametersData[i].Name;

        //worksheet.Cells[1, 1, 1, 7 + parametersData.Count].Style.Locked = true;
        using (var range = worksheet.Cells[1, 1, 1, 7 + parametersData.Count])
        {
            range.Style.Font.Bold = true;
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        var skippedRow = 0;
        for (int i = 0; i < elements.Count; i++)
        {
            var element = elements[i] as FamilyInstance;
            if (element is null)
            {
                skippedRow++;
                continue;
            }
            var location = element.Location as LocationPoint;
            if (location is null)
            {
                skippedRow++;
                continue;
            }

            worksheet.Cells[i + 2 - skippedRow, 1].Value = element.Name;
            worksheet.Cells[i + 2 - skippedRow, 2].Value = element.Symbol.FamilyName;
            worksheet.Cells[i + 2 - skippedRow, 3].Value = element.Category.Name;
            worksheet.Cells[i + 2 - skippedRow, 4].Value = location.Point.X;
            worksheet.Cells[i + 2 - skippedRow, 5].Value = location.Point.Y;
            worksheet.Cells[i + 2 - skippedRow, 6].Value = location.Point.Z;
            worksheet.Cells[i + 2 - skippedRow, 7].Value = location.Rotation;

            for (int j = 0; j < parametersData.Count; j++)
            {
                var parameterData = parametersData[j];
                var value = element.GetParameterValue(parameterData.Name);
                worksheet.Cells[i + 2 - skippedRow, 8 + j].Value = value;
            }
        }

        //worksheet.Cells[2, 1, 1 + elements.Count, 7 + parametersData.Count].Style.Locked = false;
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        //worksheet.Protection.IsProtected = true;
        //worksheet.Protection.SetPassword("1234"); // Optional password

        package.SaveAs(new FileInfo(filePath));
    }
}
