using System.Collections.Generic;
using Autodesk.Revit.DB;
using OpeningTask.Helpers;

namespace OpeningTask.Models
{
    /// <summary>
    /// Filter type for elements
    /// </summary>
    public enum FilterType
    {
        /// <summary>
        /// Selected elements
        /// </summary>
        Selected,

        /// <summary>
        /// By filter (types and parameters)
        /// </summary>
        ByFilter
    }

    /// <summary>
    /// Filter settings for section
    /// </summary>
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

        /// <summary>
        /// Filter type
        /// </summary>
        public FilterType FilterType
        {
            get => _filterType;
            set => SetProperty(ref _filterType, value);
        }

        /// <summary>
        /// Is filter enabled
        /// </summary>
        public bool IsFilterEnabled
        {
            get => _isFilterEnabled;
            set => SetProperty(ref _isFilterEnabled, value);
        }

        /// <summary>
        /// Filtered elements count
        /// </summary>
        public int ElementCount
        {
            get => _elementCount;
            set => SetProperty(ref _elementCount, value);
        }

        /// <summary>
        /// Selected type IDs
        /// </summary>
        public List<ElementId> SelectedTypeIds
        {
            get => _selectedTypeIds;
            set => SetProperty(ref _selectedTypeIds, value);
        }

        /// <summary>
        /// Selected type names (for filtering across different linked models)
        /// </summary>
        public List<string> SelectedTypeNames
        {
            get => _selectedTypeNames;
            set => SetProperty(ref _selectedTypeNames, value);
        }

        /// <summary>
        /// Selected parameter values (key - parameter name, value - list of values)
        /// </summary>
        public Dictionary<string, List<string>> SelectedParameterValues
        {
            get => _selectedParameterValues;
            set => SetProperty(ref _selectedParameterValues, value);
        }

        /// <summary>
        /// List of filtered element IDs (for current document)
        /// </summary>
        public List<ElementId> FilteredElementIds
        {
            get => _filteredElementIds;
            set => SetProperty(ref _filteredElementIds, value);
        }

        /// <summary>
        /// List of selected elements from linked models
        /// </summary>
        public List<LinkedElementInfo> SelectedLinkedElements
        {
            get => _selectedLinkedElements;
            set => SetProperty(ref _selectedLinkedElements, value);
        }

        /// <summary>
        /// Reset filter settings
        /// </summary>
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
