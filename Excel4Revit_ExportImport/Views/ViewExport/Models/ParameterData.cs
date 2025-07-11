using Excel4Revit_ExportImport.Models;

namespace Excel4Revit_ExportImport.Views.ViewExport.Models;
public class ParameterData : BaseModel
{
    public string Name { get; set; }
    public ElementId ElementId { get; set; }
    public StorageType StorageType { get; set; }
}
