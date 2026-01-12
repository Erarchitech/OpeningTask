using System.Collections.Generic;
using Autodesk.Revit.DB;
using OpeningTask.Helpers;

namespace OpeningTask.Models
{
    // Тип фильтрации элементов
    public enum FilterType
    {
        // Выбранные элементы
        Selected,

        // По фильтру (типы и параметры)
        ByFilter
    }

    // Настройки фильтра
    public class FilterSettings : ViewModels.BaseViewModel
    {
        private FilterType _filterType = FilterType.Selected;
        private bool _isFilterEnabled;
        private int _elementCount;
        private List<ElementId> _selectedTypeIds = new List<ElementId>();
        private List<string> _selectedTypeNames = new List<string>();
        private Dictionary<string, List<string>> _selectedParameterValues = new Dictionary<string, List<string>>();
        private List<ElementId> _filteredElementIds = new List<ElementId>();
        private List<LinkedElementInfo> _selectedLinkedElements = new List<LinkedElementInfo>();

        // Тип фильтра
        public FilterType FilterType
        {
            get => _filterType;
            set => SetProperty(ref _filterType, value);
        }

        // Включён ли фильтр
        public bool IsFilterEnabled
        {
            get => _isFilterEnabled;
            set => SetProperty(ref _isFilterEnabled, value);
        }

        // Количество отфильтрованных элементов
        public int ElementCount
        {
            get => _elementCount;
            set => SetProperty(ref _elementCount, value);
        }

        // ID выбранных типов
        public List<ElementId> SelectedTypeIds
        {
            get => _selectedTypeIds;
            set => SetProperty(ref _selectedTypeIds, value);
        }

        // Имена выбранных типов (для фильтрации по разным связанным моделям)
        public List<string> SelectedTypeNames
        {
            get => _selectedTypeNames;
            set => SetProperty(ref _selectedTypeNames, value);
        }

        // Выбранные значения параметров
        public Dictionary<string, List<string>> SelectedParameterValues
        {
            get => _selectedParameterValues;
            set => SetProperty(ref _selectedParameterValues, value);
        }

        // Список ID отфильтрованных элементов
        public List<ElementId> FilteredElementIds
        {
            get => _filteredElementIds;
            set => SetProperty(ref _filteredElementIds, value);
        }

        // Список выбранных элементов из связанных моделей
        public List<LinkedElementInfo> SelectedLinkedElements
        {
            get => _selectedLinkedElements;
            set => SetProperty(ref _selectedLinkedElements, value);
        }

        // Сброс настроек фильтра
        public void Reset()
        {
            FilterType = FilterType.Selected;
            IsFilterEnabled = false;
            ElementCount = 0;
            SelectedTypeIds.Clear();
            SelectedTypeNames.Clear();
            SelectedParameterValues.Clear();
            FilteredElementIds.Clear();
            SelectedLinkedElements.Clear();
        }
    }
}
