using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OpeningTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpeningTask.Helpers
{
    // Вспомогательные функции для работы с Revit
    public static class Functions
    {
        // Категории MEP систем
        public static readonly BuiltInCategory[] MepCategories = new[]
        {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_FlexDuctCurves
        };

        // Получение всех связанных моделей из документа
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
                    // Пропуск незагруженных связей
                }
            }

            return result;
        }

        // Фильтрация элементов по типам и параметрам
        public static List<Element> FilterElements(Document doc, FilterSettings settings, BuiltInCategory[] categories)
        {
            var result = new List<Element>();

            if (!settings.IsFilterEnabled) return result;

            var multiCategoryFilter = new ElementMulticategoryFilter(categories);
            var elements = new FilteredElementCollector(doc)
                .WherePasses(multiCategoryFilter)
                .WhereElementIsNotElementType()
                .ToList();

            // Фильтрация по именам типов (работает для разных связанных моделей)
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

            // Фильтрация по параметрам
            if (settings.SelectedParameterValues.Any())
            {
                elements = elements.Where(e =>
                {
                    foreach (var kvp in settings.SelectedParameterValues)
                    {
                        var param = e.LookupParameter(kvp.Key);
                        if (param == null) return false;

                        var value = param.GetStringValue();
                        if (!kvp.Value.Contains(value)) return false;
                    }
                    return true;
                }).ToList();
            }

            return elements;
        }

        // Получение имени категории
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

        // Интерактивный выбор элементов из связанных моделей
        public static List<LinkedElementInfo> SelectElementsFromLinkedModels(
            UIDocument uiDoc, 
            IEnumerable<LinkedModelInfo> selectedLinkedModels,
            BuiltInCategory[] categories)
        {
            var result = new List<LinkedElementInfo>();
            
            try
            {
                var selection = uiDoc.Selection;
                
                // Создание фильтра для связанных элементов
                var linkedModelFilter = new LinkedElementSelectionFilter(uiDoc.Document, selectedLinkedModels, categories);
                
                // Запрос выбора элементов из связанных моделей
                var references = selection.PickObjects(
                    ObjectType.LinkedElement,
                    linkedModelFilter,
                    "Select elements from linked models (press Esc to finish)");

                foreach (var reference in references)
                {
                    // Получение экземпляра связи
                    var linkInstance = uiDoc.Document.GetElement(reference.ElementId) as RevitLinkInstance;
                    if (linkInstance == null) continue;

                    var linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null) continue;

                    // Получение связанного элемента
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
                // Пользователь отменил выбор
            }
            catch (Exception)
            {
                // Другая ошибка
            }

            return result;
        }

    }

    // Информация об элементе из связанной модели
    public class LinkedElementInfo
    {
        public RevitLinkInstance LinkInstance { get; set; }
        public ElementId LinkedElementId { get; set; }
        public Element LinkedElement { get; set; }
        public Document LinkedDocument { get; set; }
    }

    // Фильтр выбора для связанных элементов
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
            // Разрешаем только выбранные экземпляры связей
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

                // Проверка, что экземпляр связи разрешён
                if (!_allowedLinkIds.Contains(reference.ElementId))
                    return false;

                // Проверка категории связанного элемента
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
