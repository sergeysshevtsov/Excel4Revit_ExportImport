using Autodesk.Revit.DB;
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

        worksheet.Cells[1, 1].Value = "GUID";
        worksheet.Cells[1, 2].Value = "Name";
        worksheet.Cells[1, 3].Value = "Family type";
        worksheet.Cells[1, 4].Value = "Category";
        worksheet.Cells[1, 5].Value = "Location X";
        worksheet.Cells[1, 6].Value = "Location Y";
        worksheet.Cells[1, 7].Value = "Location Z";
        worksheet.Cells[1, 8].Value = "Rotation";

        for (int i = 0; i < parametersData.Count; i++)
            worksheet.Cells[1, 9 + i].Value = parametersData[i].Name;

        //worksheet.Cells[1, 1, 1, 7 + parametersData.Count].Style.Locked = true;
        using (var range = worksheet.Cells[1, 1, 1, 8 + parametersData.Count])
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

            worksheet.Cells[i + 2 - skippedRow, 1].Value = element.UniqueId;
            worksheet.Cells[i + 2 - skippedRow, 2].Value = element.Name;
            worksheet.Cells[i + 2 - skippedRow, 3].Value = element.Symbol.FamilyName;
            worksheet.Cells[i + 2 - skippedRow, 4].Value = element.Category.Name;
            worksheet.Cells[i + 2 - skippedRow, 5].Value = location.Point.X;
            worksheet.Cells[i + 2 - skippedRow, 6].Value = location.Point.Y;
            worksheet.Cells[i + 2 - skippedRow, 7].Value = location.Point.Z;
            worksheet.Cells[i + 2 - skippedRow, 8].Value = location.Rotation;

            for (int j = 0; j < parametersData.Count; j++)
            {
                var parameterData = parametersData[j];
                var value = element.GetParameterValue(parameterData.Name);
                worksheet.Cells[i + 2 - skippedRow, 9 + j].Value = value;
            }
        }

        //worksheet.Cells[2, 1, 1 + elements.Count, 7 + parametersData.Count].Style.Locked = false;
        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        worksheet.Column(1).Hidden = true;
        //worksheet.Protection.IsProtected = true;
        //worksheet.Protection.SetPassword("1234"); // Optional password

        package.SaveAs(new FileInfo(filePath));
    }

    public static void Import(Document document, string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Excel4Revit_ExportImport");
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        int rows = worksheet.Dimension.Rows;
        int cols = worksheet.Dimension.Columns;

        for (int row = 2; row <= rows; row++)
        {
            var uniqueId = worksheet.Cells[row, 1].Value?.ToString();

            var X = double.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out double x) ? x : double.NaN;
            var Y = double.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out double y) ? y : double.NaN;
            var Z = double.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out double z) ? z : double.NaN;

            var R = double.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out double r) ? r : double.NaN;

            List<string> parameterNames = [];
            for (int p = 9; p <= cols; p++)
                parameterNames.Add(worksheet.Cells[1, p].Value?.ToString() ?? string.Empty);

            if (!string.IsNullOrEmpty(uniqueId))
            {
                if (document.GetElement(uniqueId) is FamilyInstance familyInstance)
                {
                    if (X != double.NaN && Y != double.NaN && Z != double.NaN)
                        familyInstance.MoveTo(new XYZ(X, Y, Z));

                    if (R != double.NaN)
                        familyInstance.RotateTo(R);


                    List<string> parameterValues = [];
                    for (int p = 9; p <= cols; p++)
                    {
                        var value = worksheet.Cells[row, p].Value?.ToString();
                        if (!string.IsNullOrEmpty(value))
                            parameterValues.Add(value);
                        else
                            parameterValues.Add(string.Empty);
                    }

                    familyInstance.SetParametersByName(parameterNames, parameterValues);
                }
            }
            else
            {
                var familyName = worksheet.Cells[row, 2].Value?.ToString();
                var familyType = worksheet.Cells[row, 3].Value?.ToString();

                if (!string.IsNullOrEmpty(familyType) && !string.IsNullOrEmpty(familyName))
                {
                    var location = XYZ.Zero;
                    if (X != double.NaN && Y != double.NaN && Z != double.NaN)
                        location = new XYZ(X, Y, Z);

                    var element = Elements.PlaceFamilyInstance(document, familyName, familyType, location, document.ActiveView.GenLevel);

                    List<string> parameterValues = [];
                    for (int p = 9; p <= cols; p++)
                    {
                        var value = worksheet.Cells[row, p].Value?.ToString();
                        if (!string.IsNullOrEmpty(value))
                            parameterValues.Add(value);
                        else
                            parameterValues.Add(string.Empty);
                    }

                    element.SetParametersByName(parameterNames, parameterValues);
                }
            }
        }
    }
}
