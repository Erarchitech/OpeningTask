using Autodesk.Revit.DB;
using System;

namespace OpeningTask.Helpers
{
    // Методы расширения для Revit API
    public static class Extensions
    {
        // Получение строкового значения параметра
        public static string GetStringValue(this Parameter param)
        {
            if (param == null) return null;

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString();
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return Math.Round(param.AsDouble(), 2).ToString();
                case StorageType.ElementId:
                    var id = param.AsElementId();
                    if (id == ElementId.InvalidElementId) return null;
                    var element = param.Element?.Document?.GetElement(id);
                    return element?.Name;
                default:
                    return null;
            }
        }
    }
}
