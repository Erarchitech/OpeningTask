using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpeningTask.Helpers
{
    // Методы расширения для Revit API
    public static class Extensions
    {
        // Получение значения параметра как строки
        public static string GetParameterValueAsString(this Element element, string parameterName)
        {
            var parameter = element.LookupParameter(parameterName);
            if (parameter == null) return null;

            switch (parameter.StorageType)
            {
                case StorageType.String:
                    return parameter.AsString();
                case StorageType.Integer:
                    return parameter.AsInteger().ToString();
                case StorageType.Double:
                    return parameter.AsDouble().ToString("F2");
                case StorageType.ElementId:
                    var id = parameter.AsElementId();
                    if (id == ElementId.InvalidElementId) return null;
                    var doc = element.Document;
                    var linkedElement = doc.GetElement(id);
                    return linkedElement?.Name;
                default:
                    return null;
            }
        }

        // Получение всех параметров элемента
        public static IEnumerable<Parameter> GetAllParameters(this Element element)
        {
            var parameters = new List<Parameter>();
            foreach (Parameter param in element.Parameters)
            {
                parameters.Add(param);
            }
            return parameters;
        }

        // Получение имени типа элемента
        public static string GetTypeName(this Element element)
        {
            var typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId) return "Unknown type";

            var type = element.Document.GetElement(typeId);
            return type?.Name ?? "Unknown type";
        }

        // Получение имени категории
        public static string GetCategoryName(this Element element)
        {
            return element.Category?.Name ?? "No category";
        }

        // Проверка, является ли элемент MEP элементом
        public static bool IsMepElement(this Element element)
        {
            if (element?.Category == null) return false;

            var mepCategories = new[]
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_DuctAccessory
            };

            // Используем свойство Value для Revit 2024 (IntegerValue устарело)
            var categoryId = element.Category.Id.Value;
            return mepCategories.Any(c => (long)c == categoryId);
        }

        // Проверка, является ли категория MEP категорией
        public static bool IsMepCategory(this BuiltInCategory category)
        {
            var mepCategories = new[]
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_DuctAccessory
            };

            return mepCategories.Contains(category);
        }

        // Безопасное получение документа из связи
        public static Document GetLinkDocumentSafe(this RevitLinkInstance linkInstance)
        {
            try
            {
                return linkInstance.GetLinkDocument();
            }
            catch
            {
                return null;
            }
        }
    }
}
