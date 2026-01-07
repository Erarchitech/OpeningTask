using Autodesk.Revit.DB;
using OpeningTask.Helpers;
using OpeningTask.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace OpeningTask.ViewModels
{
    /// <summary>
    /// ViewModel for filter window
    /// </summary>
    public class FilterWindowViewModel : BaseViewModel
    {
        private readonly List<LinkedModelInfo> _linkedModels;
        private readonly BuiltInCategory[] _categories;
        private readonly List<LinkedElementInfo> _preSelectedElements;
        private Window _window;

        private ObservableCollection<TreeNode> _typeNodes;
        private ObservableCollection<TreeNode> _parameterNodes;
        private int _selectedTabIndex;
        private string _title;
        private FilterSettings _filterSettings;

        public ICommand ApplyFilterCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand SelectAllTypesCommand { get; }
        public ICommand DeselectAllTypesCommand { get; }
        public ICommand SelectAllParametersCommand { get; }
        public ICommand DeselectAllParametersCommand { get; }

        public FilterWindowViewModel(
            IEnumerable<LinkedModelInfo> linkedModels,
            BuiltInCategory[] categories,
            FilterSettings existingSettings,
            string title,
            List<LinkedElementInfo> preSelectedElements = null)
        {
            _linkedModels = linkedModels?.ToList() ?? new List<LinkedModelInfo>();
            _categories = categories ?? throw new ArgumentNullException(nameof(categories));
            _title = title;
            _filterSettings = existingSettings ?? new FilterSettings();
            _preSelectedElements = preSelectedElements;

            _typeNodes = new ObservableCollection<TreeNode>();
            _parameterNodes = new ObservableCollection<TreeNode>();

            // Initialize commands
            ApplyFilterCommand = new RelayCommand(ApplyFilter);
            CancelCommand = new RelayCommand(Cancel);
            SelectAllTypesCommand = new RelayCommand(SelectAllTypes);
            DeselectAllTypesCommand = new RelayCommand(DeselectAllTypes);
            SelectAllParametersCommand = new RelayCommand(SelectAllParameters);
            DeselectAllParametersCommand = new RelayCommand(DeselectAllParameters);

            // Load data
            LoadTypes();
        }

        public void SetWindow(Window window)
        {
            _window = window;
        }

        #region Properties

        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        public ObservableCollection<TreeNode> TypeNodes
        {
            get => _typeNodes;
            set => SetProperty(ref _typeNodes, value);
        }

        public ObservableCollection<TreeNode> ParameterNodes
        {
            get => _parameterNodes;
            set => SetProperty(ref _parameterNodes, value);
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set
            {
                if (SetProperty(ref _selectedTabIndex, value))
                {
                    if (value == 1) // Parameters tab
                    {
                        LoadParameters();
                    }
                }
            }
        }

        public FilterSettings FilterSettings
        {
            get => _filterSettings;
            set => SetProperty(ref _filterSettings, value);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Get valid linked models (loaded ones)
        /// </summary>
        private IEnumerable<LinkedModelInfo> GetValidLinkedModels()
        {
            return _linkedModels.Where(m => m.IsLoaded);
        }

        /// <summary>
        /// Get elements to filter - either pre-selected or from linked models
        /// </summary>
        private List<Element> GetElementsToFilter()
        {
            var elements = new List<Element>();

            // If we have pre-selected elements, use only those
            if (_preSelectedElements != null && _preSelectedElements.Any())
            {
                elements.AddRange(_preSelectedElements
                    .Where(e => e.LinkedElement != null)
                    .Select(e => e.LinkedElement));
            }
            else
            {
                // Otherwise, get all elements from linked models
                var validModels = GetValidLinkedModels().ToList();

                foreach (var linkedModel in validModels)
                {
                    var doc = linkedModel.LinkedDocument;
                    if (doc == null) continue;

                    try
                    {
                        var multiCategoryFilter = new ElementMulticategoryFilter(_categories);
                        var modelElements = new FilteredElementCollector(doc)
                            .WherePasses(multiCategoryFilter)
                            .WhereElementIsNotElementType()
                            .ToList();

                        elements.AddRange(modelElements);
                    }
                    catch (Exception)
                    {
                        // Skip if error
                    }
                }
            }

            return elements;
        }

        /// <summary>
        /// Load element types
        /// </summary>
        private void LoadTypes()
        {
            _typeNodes.Clear();

            // Get types from pre-selected elements or from linked models
            if (_preSelectedElements != null && _preSelectedElements.Any())
            {
                LoadTypesFromPreSelectedElements();
            }
            else
            {
                LoadTypesFromLinkedModels();
            }

            // Update parent node states
            foreach (var categoryNode in _typeNodes)
            {
                UpdateCategoryNodeState(categoryNode);
            }
        }

        /// <summary>
        /// Load types from pre-selected elements
        /// </summary>
        private void LoadTypesFromPreSelectedElements()
        {
            var typesByCategory = new Dictionary<string, Dictionary<string, ElementId>>();

            foreach (var elementInfo in _preSelectedElements.Where(e => e.LinkedElement != null))
            {
                var element = elementInfo.LinkedElement;
                var doc = elementInfo.LinkedDocument;
                
                if (element.Category == null) continue;

                var categoryName = element.Category.Name;
                var typeId = element.GetTypeId();
                
                if (typeId == ElementId.InvalidElementId) continue;

                var typeElement = doc.GetElement(typeId);
                if (typeElement == null) continue;

                var typeName = typeElement.Name;

                if (!typesByCategory.ContainsKey(categoryName))
                {
                    typesByCategory[categoryName] = new Dictionary<string, ElementId>();
                }

                if (!typesByCategory[categoryName].ContainsKey(typeName))
                {
                    typesByCategory[categoryName][typeName] = typeId;
                }
            }

            // Create tree nodes
            foreach (var category in typesByCategory.OrderBy(c => c.Key))
            {
                var categoryNode = new TreeNode(category.Key) { IsExpanded = true };
                _typeNodes.Add(categoryNode);

                foreach (var type in category.Value.OrderBy(t => t.Key))
                {
                    var typeNode = categoryNode.AddChild(type.Key, type.Value);

                    // Restore previous selection
                    if (_filterSettings.SelectedTypeIds.Any(id => id.Value == type.Value.Value))
                    {
                        typeNode.SetCheckedSilent(true);
                    }
                }
            }
        }

        /// <summary>
        /// Load types from linked models
        /// </summary>
        private void LoadTypesFromLinkedModels()
        {
            var validModels = GetValidLinkedModels().ToList();

            if (!validModels.Any())
            {
                return;
            }

            foreach (var linkedModel in validModels)
            {
                var doc = linkedModel.LinkedDocument;
                if (doc == null) continue;

                foreach (var category in _categories)
                {
                    var categoryName = Functions.GetCategoryName(doc, category);

                    // Check if category already exists in tree
                    var categoryNode = _typeNodes.FirstOrDefault(n => n.Name == categoryName);
                    if (categoryNode == null)
                    {
                        categoryNode = new TreeNode(categoryName) { IsExpanded = true };
                        _typeNodes.Add(categoryNode);
                    }

                    // Get types for category
                    try
                    {
                        var types = new FilteredElementCollector(doc)
                            .OfCategory(category)
                            .WhereElementIsElementType()
                            .Cast<ElementType>()
                            .OrderBy(t => t.Name)
                            .ToList();

                        foreach (var type in types)
                        {
                            // Check if type already exists
                            if (!categoryNode.Children.Any(c => c.Name == type.Name))
                            {
                                var typeNode = categoryNode.AddChild(type.Name, type.Id);

                                // Restore previous selection
                                if (_filterSettings.SelectedTypeIds.Any(id => id.Value == type.Id.Value))
                                {
                                    typeNode.SetCheckedSilent(true);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Skip if error getting types
                    }
                }
            }
        }

        /// <summary>
        /// Load parameters for selected types
        /// </summary>
        private void LoadParameters()
        {
            _parameterNodes.Clear();

            // Get selected types
            var selectedTypeIds = GetSelectedTypeIds();

            // Get elements to analyze
            var elements = GetElementsToFilter();

            // Filter by selected types if any
            if (selectedTypeIds.Any())
            {
                var typeIdValues = selectedTypeIds.Select(id => id.Value).ToHashSet();
                elements = elements.Where(e => typeIdValues.Contains(e.GetTypeId().Value)).ToList();
            }

            // Collect parameters and their values
            var parameterValues = new Dictionary<string, HashSet<string>>();

            foreach (var element in elements.Take(1000)) // Limit for performance
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param.Definition == null || param.IsReadOnly) continue;
                    if (param.StorageType == StorageType.None) continue;

                    var paramName = param.Definition.Name;
                    var value = GetParameterStringValue(param);

                    if (string.IsNullOrWhiteSpace(value)) continue;

                    if (!parameterValues.ContainsKey(paramName))
                    {
                        parameterValues[paramName] = new HashSet<string>();
                    }

                    parameterValues[paramName].Add(value);
                }
            }

            // Create parameter tree nodes
            foreach (var param in parameterValues.OrderBy(p => p.Key))
            {
                var paramNode = new TreeNode(param.Key) { IsExpanded = false };

                foreach (var value in param.Value.OrderBy(v => v))
                {
                    var valueNode = paramNode.AddChild(value, value);

                    // Restore previous selection
                    if (_filterSettings.SelectedParameterValues.TryGetValue(param.Key, out var selectedValues))
                    {
                        if (selectedValues.Contains(value))
                        {
                            valueNode.SetCheckedSilent(true);
                        }
                    }
                }

                // Update parent node state
                UpdateCategoryNodeState(paramNode);

                _parameterNodes.Add(paramNode);
            }
        }

        /// <summary>
        /// Get selected type IDs
        /// </summary>
        private List<ElementId> GetSelectedTypeIds()
        {
            return _typeNodes
                .SelectMany(c => c.Children)
                .Where(t => t.IsChecked == true && t.Tag is ElementId)
                .Select(t => (ElementId)t.Tag)
                .ToList();
        }

        /// <summary>
        /// Get selected parameter values
        /// </summary>
        private Dictionary<string, List<string>> GetSelectedParameterValues()
        {
            var result = new Dictionary<string, List<string>>();

            foreach (var paramNode in _parameterNodes)
            {
                var selectedValues = paramNode.Children
                    .Where(v => v.IsChecked == true && v.Tag is string)
                    .Select(v => (string)v.Tag)
                    .ToList();

                if (selectedValues.Any())
                {
                    result[paramNode.Name] = selectedValues;
                }
            }

            return result;
        }

        /// <summary>
        /// Count filtered elements
        /// </summary>
        private int CountFilteredElements()
        {
            var selectedTypeIds = GetSelectedTypeIds();
            var selectedParamValues = GetSelectedParameterValues();

            var elements = GetElementsToFilter();

            // Filter by types
            if (selectedTypeIds.Any())
            {
                var typeIdValues = selectedTypeIds.Select(id => id.Value).ToHashSet();
                elements = elements
                    .Where(e => typeIdValues.Contains(e.GetTypeId().Value))
                    .ToList();
            }

            // Filter by parameters
            if (selectedParamValues.Any())
            {
                elements = elements.Where(e =>
                {
                    foreach (var kvp in selectedParamValues)
                    {
                        var param = e.LookupParameter(kvp.Key);
                        if (param == null) continue;

                        var value = GetParameterStringValue(param);
                        if (value != null && kvp.Value.Contains(value))
                        {
                            return true;
                        }
                    }
                    return false;
                }).ToList();
            }

            return elements.Count;
        }

        /// <summary>
        /// Get parameter string value
        /// </summary>
        private string GetParameterStringValue(Parameter param)
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
        /// Update category node state based on children
        /// </summary>
        private void UpdateCategoryNodeState(TreeNode categoryNode)
        {
            if (!categoryNode.Children.Any()) return;

            var checkedCount = categoryNode.Children.Count(c => c.IsChecked == true);
            var uncheckedCount = categoryNode.Children.Count(c => c.IsChecked == false);

            if (checkedCount == categoryNode.Children.Count)
            {
                categoryNode.SetCheckedSilent(true);
            }
            else if (uncheckedCount == categoryNode.Children.Count)
            {
                categoryNode.SetCheckedSilent(false);
            }
            else
            {
                categoryNode.SetCheckedSilent(null);
            }
        }

        /// <summary>
        /// Apply filter
        /// </summary>
        private void ApplyFilter()
        {
            _filterSettings.SelectedTypeIds = GetSelectedTypeIds();
            _filterSettings.SelectedParameterValues = GetSelectedParameterValues();
            _filterSettings.ElementCount = CountFilteredElements();
            _filterSettings.IsFilterEnabled = true;

            if (_window != null)
            {
                _window.DialogResult = true;
                _window.Close();
            }
        }

        /// <summary>
        /// Cancel
        /// </summary>
        private void Cancel()
        {
            if (_window != null)
            {
                _window.DialogResult = false;
                _window.Close();
            }
        }

        /// <summary>
        /// Select all types
        /// </summary>
        private void SelectAllTypes()
        {
            foreach (var categoryNode in _typeNodes)
            {
                categoryNode.IsChecked = true;
            }
        }

        /// <summary>
        /// Deselect all types
        /// </summary>
        private void DeselectAllTypes()
        {
            foreach (var categoryNode in _typeNodes)
            {
                categoryNode.IsChecked = false;
            }
        }

        /// <summary>
        /// Select all parameters
        /// </summary>
        private void SelectAllParameters()
        {
            foreach (var paramNode in _parameterNodes)
            {
                paramNode.IsChecked = true;
            }
        }

        /// <summary>
        /// Deselect all parameters
        /// </summary>
        private void DeselectAllParameters()
        {
            foreach (var paramNode in _parameterNodes)
            {
                paramNode.IsChecked = false;
            }
        }

        #endregion
    }
}
