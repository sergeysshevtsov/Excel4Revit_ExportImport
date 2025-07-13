using System.Windows.Media.Imaging;

namespace Excel4Revit_ExportImport.RevitExtensions;
public static class MoveExtension
{
    public static void MoveTo(this Element element, XYZ newLocation)
    {
        var localPoint = element.Location as LocationPoint;
        var point = localPoint?.Point;
        ElementTransformUtils.MoveElement(element.Document, element.Id, new(newLocation.X - point.X, newLocation.Y - point.Y, newLocation.Z - point.Z));
    }

    public static void RotateTo(this Element element, double rotation)
    {
        if (element.Location is not LocationPoint localPoint)
            return;
       
        var bb = element.get_BoundingBox(null);
        if (bb == null)
            return;

        XYZ pnt = new((bb.Min.X + bb.Max.X) / 2, (bb.Min.Y + bb.Max.Y) / 2, bb.Min.Z);
        var axis = Line.CreateBound(pnt, new XYZ(pnt.X, pnt.Y, pnt.Z + 10));

        var currentRotation = localPoint.Rotation;
        ElementTransformUtils.RotateElement(element.Document, element.Id, axis, (-1) * currentRotation);
        element.Document.Regenerate();
        ElementTransformUtils.RotateElement(element.Document, element.Id, axis, rotation);
    }
}
