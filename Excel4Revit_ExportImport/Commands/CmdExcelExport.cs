using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Excel4Revit_ExportImport.RevitExtensions;
using Excel4Revit_ExportImport.Views.ViewExport;
using Nice3point.Revit.Toolkit.External;
using System.Diagnostics;
using System.IO;
using ZLinq;

namespace Excel4Revit_ExportImport.Commands;

[UsedImplicitly]
[Transaction(TransactionMode.Manual)]
public class CmdExcelExport : ExternalCommand
{
    public override void Execute()
    {
        var uiapp = ExternalCommandData.Application;
        var uidoc = uiapp.ActiveUIDocument;
        var application = uiapp.Application;
        var document = uidoc.Document;

        var set = uidoc.Selection.GetElementIds();

        if (set.Count == 0)
        {
            TaskDialog.Show("Model Exporter", "No elements selected for export.");
            return;
        }

        var elements = new FilteredElementCollector(document, set)
            .WhereElementIsNotElementType()
            .WhereElementIsViewIndependent()

            .ToList();

        var dinstinctElements = elements
            .AsValueEnumerable()
            .GroupBy(x => x.Name)
            .Select(g => g.First())
            .ToList();

        var parametersData = dinstinctElements.GetParametersData();

        var dialog = new Views.ViewExport.WindowExport(elements, parametersData);
        new System.Windows.Interop.WindowInteropHelper(dialog)
        {
            Owner = Autodesk.Windows.ComponentManager.ApplicationWindow
        };
        dialog.ShowDialog();

        if (!(dialog.DataContext as WindowExportDataContext).IsExported)
            return;

        TaskDialog downloadDialog = new("Excel file is ready")
        {
            MainContent = $"Open Excel file",
            AllowCancellation = false
        };
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                   "Open Excel file",
                   "This option will open Excel file.");
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                   "Navigate to Excel file",
                   "This option will open Excel file directory.");
        downloadDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                   "Close");

        var pathToFile = Path.Combine(Path.GetTempPath(), "ExcelExport.xlsx");
        switch (downloadDialog.Show())
        {
            case TaskDialogResult.CommandLink1:
                try
                {
                    Process.Start(new ProcessStartInfo(pathToFile) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Error", $"Failed to open file: {ex.Message}");
                }
                break;

            case TaskDialogResult.CommandLink2:
                Process.Start("explorer.exe", $"/select,\"{pathToFile}\"");
                break;

            default:
                break;
        }
    }
}