using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;


namespace Excel4Revit_ExportImport.RevitExtensions;
public class Elements
{
    public static Element PlaceFamilyInstance(Document document, string familyName, string typeName, XYZ location, Level level)
    {
        FamilySymbol symbol = new FilteredElementCollector(document)
           .OfClass(typeof(FamilySymbol))
           .Cast<FamilySymbol>()
           .FirstOrDefault(fs =>
               fs.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
               fs.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));

        if (symbol == null)
        {
            TaskDialog.Show("Error", "Family type not found.");
            return null;
        }

        if (!symbol.IsActive)
        {
            symbol.Activate();
            document.Regenerate();
        }

        FamilyInstance instance = document.Create.NewFamilyInstance(
            location,
            symbol,
            level,
            StructuralType.NonStructural
        );

        return instance;
    }
}
