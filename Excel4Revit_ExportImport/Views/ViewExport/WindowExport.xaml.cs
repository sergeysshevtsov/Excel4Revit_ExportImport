using Excel4Revit_ExportImport.Views.ViewExport.Models;
using System.Windows;

namespace Excel4Revit_ExportImport.Views.ViewExport;
public partial class WindowExport : Window
{
    public WindowExport(List<Element> elements, List<ParameterData> parametersData)
    {
        InitializeComponent();
        DataContext = new WindowExportDataContext(elements, parametersData);
    }
}
