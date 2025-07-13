using Excel4Revit_ExportImport.Views.ViewExport.Models;

namespace Excel4Revit_ExportImport.RevitExtensions;
public static class ParametersExstension
{
    public static List<ParameterData> GetParametersData(this List<Element> elements)
    {
        var result = new List<ParameterData>();
        if (elements is null || elements.Count == 0)
            return result;

        foreach (var element in elements)
        {
            var parameterData = new ParameterData();
            if (element is null)
                continue;

            if (element is FamilyInstance familyInstance)
            {
                foreach (var p in element.GetOrderedParameters())
                {
                    if (p is Parameter parameter)
                    {
                        if (parameter == null)
                            continue;

                        if (parameter.IsReadOnly)
                            continue;

                        if (result.AsEnumerable().FirstOrDefault(x => x.Name == parameter.Definition.Name) == null)
                        {
                            result.Add(new ParameterData
                            {
                                Name = parameter.Definition.Name,
                                ElementId = parameter.Id,
                                StorageType = parameter.StorageType
                            });
                        }
                    }
                }
            }
        }

        return result;
    }

    public static string GetParameterValue(this Element element, string parameterName)
    {
        var result = string.Empty;
        Parameter parameter = element.FindParameter(parameterName);
        if (parameter == null)
            return result;

        switch (parameter.StorageType)
        {
            case StorageType.String:
                result = parameter.AsString();
                break;
            case StorageType.Double:
                result = parameter.AsDouble().ToString();
                break;
            case StorageType.Integer:
                result = parameter.AsInteger().ToString();
                break;
            case StorageType.ElementId:
                result = parameter.AsValueString() ?? parameter.AsString();
                break;
        }

        return result;
    }

    public static void SetParametersByName(this Element element, List<string> parameterNames, List<string> parameterValues)
    {
        for (int i = 0; i < parameterNames.Count; i++)
        {
            var parameterName = parameterNames[i];
            var parameterValue = i < parameterValues.Count ? parameterValues[i] : null;
            SetParameterByName(element, parameterName, parameterValue);
        }
    }

    private static bool SetParameterByName(Element element, string parameterName, string parameterValue)
    {
        if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(parameterValue))
            return false;
        Parameter parameter = element.FindParameter(parameterName);
        if (parameter == null)
            return false;
        switch (parameter.StorageType)
        {
            case StorageType.String:
                return parameter.Set(parameterValue);
            case StorageType.Double:
                if (double.TryParse(parameterValue, out double doubleValue))
                    return parameter.Set(doubleValue);
                break;
            case StorageType.Integer:
                if (int.TryParse(parameterValue, out int intValue))
                    return parameter.Set(intValue);
                break;
            case StorageType.ElementId:
                if (int.TryParse(parameterValue, out int elementIdValue))
                    return parameter.Set(new ElementId(elementIdValue));
                break;
        }
        return false;
    }
}
