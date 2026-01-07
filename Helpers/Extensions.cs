using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpeningTask.Helpers
{
    /// <summary>
    /// Extension methods for Revit API
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Get parameter value as string
        /// </summary>
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

        /// <summary>
        /// Get all parameters of an element
        /// </summary>
        public static IEnumerable<Parameter> GetAllParameters(this Element element)
        {
            var parameters = new List<Parameter>();
            foreach (Parameter param in element.Parameters)
            {
                parameters.Add(param);
            }
            return parameters;
        }

        /// <summary>
        /// Get element type name
        /// </summary>
        public static string GetTypeName(this Element element)
        {
            var typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId) return "Unknown type";

            var type = element.Document.GetElement(typeId);
            return type?.Name ?? "Unknown type";
        }

        /// <summary>
        /// Get category name
        /// </summary>
        public static string GetCategoryName(this Element element)
        {
            return element.Category?.Name ?? "No category";
        }

        /// <summary>
        /// Check if element is MEP element
        /// </summary>
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

            // Use Value property for Revit 2024 (IntegerValue is obsolete)
            var categoryId = element.Category.Id.Value;
            return mepCategories.Any(c => (long)c == categoryId);
        }

        /// <summary>
        /// Check if category is MEP category
        /// </summary>
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

        /// <summary>
        /// Safe get document from link
        /// </summary>
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
