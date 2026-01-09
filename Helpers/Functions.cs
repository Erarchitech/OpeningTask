using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OpeningTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpeningTask.Helpers
{
    /// <summary>
    /// Helper functions for Revit operations
    /// </summary>
    public static class Functions
    {
        /// <summary>
        /// MEP system categories
        /// </summary>
        public static readonly BuiltInCategory[] MepCategories = new[]
        {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_FlexDuctCurves
        };

        /// <summary>
        /// Get all linked models from document
        /// </summary>
        public static List<LinkedModelInfo> GetAllLinkedModels(Document doc)
        {
            var result = new List<LinkedModelInfo>();

            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => link.GetLinkDocument() != null)
                .ToList();

            foreach (var linkInstance in linkInstances)
            {
                try
                {
                    var linkedModelInfo = new LinkedModelInfo(linkInstance);
                    if (linkedModelInfo.IsLoaded)
                    {
                        result.Add(linkedModelInfo);
                    }
                }
                catch (Exception)
                {
                    // Skip unloaded links
                }
            }

            return result;
        }

        /// <summary>
        /// Get element types by categories from document
        /// </summary>
        public static Dictionary<string, List<ElementType>> GetTypesByCategories(Document doc, BuiltInCategory[] categories)
        {
            var result = new Dictionary<string, List<ElementType>>();

            foreach (var category in categories)
            {
                var types = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsElementType()
                    .Cast<ElementType>()
                    .OrderBy(t => t.Name)
                    .ToList();

                if (types.Any())
                {
                    var categoryName = GetCategoryName(doc, category);
                    result[categoryName] = types;
                }
            }

            return result;
        }

        /// <summary>
        /// Get MEP element types from linked models
        /// </summary>
        public static Dictionary<string, List<ElementType>> GetMepTypesFromLinkedModels(IEnumerable<LinkedModelInfo> linkedModels)
        {
            var result = new Dictionary<string, List<ElementType>>();

            foreach (var linkedModel in linkedModels.Where(m => m.IsLoaded && m.IsSelected))
            {
                var doc = linkedModel.LinkedDocument;
                var typesByCategory = GetTypesByCategories(doc, MepCategories);

                foreach (var kvp in typesByCategory)
                {
                    if (result.ContainsKey(kvp.Key))
                    {
                        // Merge types avoiding duplicates by name
                        var existingNames = result[kvp.Key].Select(t => t.Name).ToHashSet();
                        var newTypes = kvp.Value.Where(t => !existingNames.Contains(t.Name));
                        result[kvp.Key].AddRange(newTypes);
                    }
                    else
                    {
                        result[kvp.Key] = new List<ElementType>(kvp.Value);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get wall types from linked models
        /// </summary>
        public static Dictionary<string, List<ElementType>> GetWallTypesFromLinkedModels(IEnumerable<LinkedModelInfo> linkedModels)
        {
            var result = new Dictionary<string, List<ElementType>>();
            var category = BuiltInCategory.OST_Walls;

            foreach (var linkedModel in linkedModels.Where(m => m.IsLoaded && m.IsSelected))
            {
                var doc = linkedModel.LinkedDocument;
                var types = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsElementType()
                    .Cast<ElementType>()
                    .OrderBy(t => t.Name)
                    .ToList();

                if (types.Any())
                {
                    var categoryName = GetCategoryName(doc, category);
                    if (result.ContainsKey(categoryName))
                    {
                        var existingNames = result[categoryName].Select(t => t.Name).ToHashSet();
                        var newTypes = types.Where(t => !existingNames.Contains(t.Name));
                        result[categoryName].AddRange(newTypes);
                    }
                    else
                    {
                        result[categoryName] = new List<ElementType>(types);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get floor types from linked models
        /// </summary>
        public static Dictionary<string, List<ElementType>> GetFloorTypesFromLinkedModels(IEnumerable<LinkedModelInfo> linkedModels)
        {
            var result = new Dictionary<string, List<ElementType>>();
            var category = BuiltInCategory.OST_Floors;

            foreach (var linkedModel in linkedModels.Where(m => m.IsLoaded && m.IsSelected))
            {
                var doc = linkedModel.LinkedDocument;
                var types = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsElementType()
                    .Cast<ElementType>()
                    .OrderBy(t => t.Name)
                    .ToList();

                if (types.Any())
                {
                    var categoryName = GetCategoryName(doc, category);
                    if (result.ContainsKey(categoryName))
                    {
                        var existingNames = result[categoryName].Select(t => t.Name).ToHashSet();
                        var newTypes = types.Where(t => !existingNames.Contains(t.Name));
                        result[categoryName].AddRange(newTypes);
                    }
                    else
                    {
                        result[categoryName] = new List<ElementType>(types);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get parameter values for selected types
        /// </summary>
        public static Dictionary<string, HashSet<string>> GetParameterValuesForTypes(Document doc, IEnumerable<ElementId> typeIds)
        {
            var result = new Dictionary<string, HashSet<string>>();

            foreach (var typeId in typeIds)
            {
                var elements = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => e.GetTypeId() == typeId)
                    .Take(100); // Limit for performance

                foreach (var element in elements)
                {
                    foreach (Parameter param in element.Parameters)
                    {
                        if (param.Definition == null || param.IsReadOnly) continue;
                        if (param.StorageType == StorageType.None) continue;

                        var paramName = param.Definition.Name;
                        var value = GetParameterStringValue(param);

                        if (string.IsNullOrWhiteSpace(value)) continue;

                        if (!result.ContainsKey(paramName))
                        {
                            result[paramName] = new HashSet<string>();
                        }

                        result[paramName].Add(value);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Filter elements by types and parameters
        /// </summary>
        public static List<Element> FilterElements(Document doc, FilterSettings settings, BuiltInCategory[] categories)
        {
            var result = new List<Element>();

            if (!settings.IsFilterEnabled) return result;

            var multiCategoryFilter = new ElementMulticategoryFilter(categories);
            var elements = new FilteredElementCollector(doc)
                .WherePasses(multiCategoryFilter)
                .WhereElementIsNotElementType()
                .ToList();

            // Filter by type names (works across different linked models)
            if (settings.SelectedTypeNames != null && settings.SelectedTypeNames.Any())
            {
                elements = elements
                    .Where(e =>
                    {
                        var typeId = e.GetTypeId();
                        if (typeId == ElementId.InvalidElementId) return false;
                        var typeElement = doc.GetElement(typeId);
                        return typeElement != null && settings.SelectedTypeNames.Contains(typeElement.Name);
                    })
                    .ToList();
            }

            // Filter by parameters
            if (settings.SelectedParameterValues.Any())
            {
                elements = elements.Where(e =>
                {
                    foreach (var kvp in settings.SelectedParameterValues)
                    {
                        var param = e.LookupParameter(kvp.Key);
                        if (param == null) return false;

                        var value = GetParameterStringValue(param);
                        if (!kvp.Value.Contains(value)) return false;
                    }
                    return true;
                }).ToList();
            }

            return elements;
        }

        /// <summary>
        /// Get category name
        /// </summary>
        public static string GetCategoryName(Document doc, BuiltInCategory category)
        {
            try
            {
                var cat = Category.GetCategory(doc, category);
                return cat?.Name ?? category.ToString();
            }
            catch
            {
                return category.ToString();
            }
        }

        /// <summary>
        /// Get parameter string value
        /// </summary>
        private static string GetParameterStringValue(Parameter param)
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

        /// <summary>
        /// Load cuboid family
        /// </summary>
        public static Family LoadCuboidFamily(Document doc, string familyPath)
        {
            Family family = null;

            using (var transaction = new Transaction(doc, "Load cuboid family"))
            {
                transaction.Start();

                if (!doc.LoadFamily(familyPath, out family))
                {
                    // Family already loaded, find it
                    var familyName = System.IO.Path.GetFileNameWithoutExtension(familyPath);
                    family = new FilteredElementCollector(doc)
                        .OfClass(typeof(Family))
                        .Cast<Family>()
                        .FirstOrDefault(f => f.Name == familyName);
                }

                transaction.Commit();
            }

            return family;
        }

        /// <summary>
        /// Select elements on view
        /// </summary>
        public static List<ElementId> SelectElementsOnView(UIDocument uiDoc, BuiltInCategory[] categories)
        {
            try
            {
                var selection = uiDoc.Selection;
                var selectedIds = selection.GetElementIds().ToList();

                if (!selectedIds.Any())
                {
                    // If nothing selected, prompt user to select
                    var reference = selection.PickObjects(
                        ObjectType.Element,
                        "Select MEP elements");

                    return reference.Select(r => r.ElementId).ToList();
                }

                // Filter already selected elements by categories
                return selectedIds
                    .Select(id => uiDoc.Document.GetElement(id))
                    .Where(e => e != null && categories.Any(c => 
                        e.Category?.Id.Value == (long)c))
                    .Select(e => e.Id)
                    .ToList();
            }
            catch
            {
                return new List<ElementId>();
            }
        }

        /// <summary>
        /// Select elements from linked models interactively
        /// </summary>
        public static List<LinkedElementInfo> SelectElementsFromLinkedModels(
            UIDocument uiDoc, 
            IEnumerable<LinkedModelInfo> selectedLinkedModels,
            BuiltInCategory[] categories)
        {
            var result = new List<LinkedElementInfo>();
            
            try
            {
                var selection = uiDoc.Selection;
                
                // Create filter for linked elements
                var linkedModelFilter = new LinkedElementSelectionFilter(uiDoc.Document, selectedLinkedModels, categories);
                
                // Prompt user to select elements from linked models
                var references = selection.PickObjects(
                    ObjectType.LinkedElement,
                    linkedModelFilter,
                    "Select elements from linked models (press Esc to finish)");

                foreach (var reference in references)
                {
                    // Get link instance
                    var linkInstance = uiDoc.Document.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance == null) continue;

                    var linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null) continue;

                    // Get linked element
                    var linkedElementId = reference.LinkedElementId;
                    var linkedElement = linkedDoc.GetElement(linkedElementId);
                    if (linkedElement == null) continue;

                    result.Add(new LinkedElementInfo
                    {
                        LinkInstance = linkInstance,
                        LinkedElementId = linkedElementId,
                        LinkedElement = linkedElement,
                        LinkedDocument = linkedDoc
                    });
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled selection - return what we have
            }
            catch (Exception)
            {
                // Other error
            }

            return result;
        }

        /// <summary>
        /// Count MEP elements in selected linked models
        /// </summary>
        public static int CountMepElementsInLinkedModels(IEnumerable<LinkedModelInfo> linkedModels, BuiltInCategory[] categories)
        {
            int count = 0;

            foreach (var linkedModel in linkedModels.Where(m => m.IsLoaded && m.IsSelected))
            {
                var doc = linkedModel.LinkedDocument;
                var multiCategoryFilter = new ElementMulticategoryFilter(categories);
                
                count += new FilteredElementCollector(doc)
                    .WherePasses(multiCategoryFilter)
                    .WhereElementIsNotElementType()
                    .GetElementCount();
            }

            return count;
        }
    }

    /// <summary>
    /// Information about element from linked model
    /// </summary>
    public class LinkedElementInfo
    {
        public RevitLinkInstance LinkInstance { get; set; }
        public ElementId LinkedElementId { get; set; }
        public Element LinkedElement { get; set; }
        public Document LinkedDocument { get; set; }
    }

    /// <summary>
    /// Selection filter for linked elements
    /// </summary>
    public class LinkedElementSelectionFilter : ISelectionFilter
    {
        private readonly Document _hostDocument;
        private readonly HashSet<ElementId> _allowedLinkIds;
        private readonly BuiltInCategory[] _categories;
        private readonly HashSet<long> _allowedCategoryIds;

        public LinkedElementSelectionFilter(Document hostDocument, IEnumerable<LinkedModelInfo> allowedLinks, BuiltInCategory[] categories)
        {
            _hostDocument = hostDocument;
            _allowedLinkIds = new HashSet<ElementId>(
                allowedLinks.Where(l => l.IsSelected && l.IsLoaded)
                           .Select(l => l.LinkInstance.Id));
            _categories = categories;

            _allowedCategoryIds = new HashSet<long>();
            if (_categories != null)
            {
                foreach (var c in _categories)
                {
                    _allowedCategoryIds.Add((long)c);
                }
            }
        }

        public bool AllowElement(Element elem)
        {
            // Allow only selected link instances
            if (elem is RevitLinkInstance linkInstance)
            {
                return _allowedLinkIds.Contains(linkInstance.Id);
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            try
            {
                if (reference == null) return false;
                if (_hostDocument == null) return false;

                // Ensure the link instance itself is allowed
                if (!_allowedLinkIds.Contains(reference.ElementId))
                    return false;

                // Check linked element category
                var linkInstance = reference.ElementId != null
                    ? _hostDocument.GetElement(reference.ElementId) as RevitLinkInstance
                    : null;

                if (linkInstance == null) return false;

                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) return false;

                var linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                var categoryId = linkedElement?.Category?.Id?.Value;

                if (categoryId == null) return false;

                return _allowedCategoryIds.Contains(categoryId.Value);
            }
            catch
            {
                return false;
            }
        }
    }
}
