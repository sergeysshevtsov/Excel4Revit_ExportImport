using Autodesk.Revit.Attributes;
using Excel4Revit_ExportImport.ExcelUtils;
using Microsoft.Win32;
using Nice3point.Revit.Toolkit.External;
using System.IO;

namespace Excel4Revit_ExportImport.Commands;

[UsedImplicitly]
[Transaction(TransactionMode.Manual)]

internal class CmdExcelImport : ExternalCommand
{
    public override void Execute()
    {
        OpenFileDialog openFileDialog = new()
        {
            Title = "Select a file",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            InitialDirectory = Path.GetTempPath()
        };

        var document = ExternalCommandData.Application.ActiveUIDocument.Document;
        using Transaction tr = new(document, "Apply changes");
        tr.Start();
        if (openFileDialog.ShowDialog() == true)
            ExcelFile.Import(document, openFileDialog.FileName);
        tr.Commit();
    }
}
