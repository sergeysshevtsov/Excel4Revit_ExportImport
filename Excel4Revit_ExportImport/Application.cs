using Excel4Revit_ExportImport.Commands;
using Nice3point.Revit.Toolkit.External;

namespace Excel4Revit_ExportImport;
[UsedImplicitly]
public class Application : ExternalApplication
{
    public override void OnStartup()
    {
        CreateRibbon();
    }

    private void CreateRibbon()
    {
        var panel = Application.CreatePanel("Export", "SHSS Tools");

        panel.AddPushButton<CmdExcelExport>("Excel\nExport")
            .SetImage("/Excel4Revit_ExportImport;component/Resources/Icons/excelExport16.png")
            .SetLargeImage("/Excel4Revit_ExportImport;component/Resources/Icons/excelExport32.png");

        panel.AddPushButton<CmdExcelImport>("Excel\nImport")
            .SetImage("/Excel4Revit_ExportImport;component/Resources/Icons/excelImport16.png")
            .SetLargeImage("/Excel4Revit_ExportImport;component/Resources/Icons/excelImport32.png");
        panel.AddSeparator();
    }
}