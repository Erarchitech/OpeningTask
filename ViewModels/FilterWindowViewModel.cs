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
    // ViewModel для окна фильтрации
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

            // Инициализация команд
            ApplyFilterCommand = new RelayCommand(ApplyFilter);
            CancelCommand = new RelayCommand(Cancel);
            SelectAllTypesCommand = new RelayCommand(SelectAllTypes);
            DeselectAllTypesCommand = new RelayCommand(DeselectAllTypes);
            SelectAllParametersCommand = new RelayCommand(SelectAllParameters);
            DeselectAllParametersCommand = new RelayCommand(DeselectAllParameters);

            // Загрузка данных
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
                    if (value == 1)
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

        // Получение доступных связанных моделей
        private IEnumerable<LinkedModelInfo> GetValidLinkedModels()
        {
            return _linkedModels.Where(m => m.IsLoaded);
        }

        // Получение элементов для фильтрации
        private List<Element> GetElementsToFilter()
        {
            var elements = new List<Element>();

            // Если есть предварительно выбранные элементы, используем только их
            if (_preSelectedElements != null && _preSelectedElements.Any())
            {
                elements.AddRange(_preSelectedElements
                    .Where(e => e.LinkedElement != null)
                    .Select(e => e.LinkedElement));
            }
            else
            {
                // Иначе получаем все элементы из связанных моделей
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
                        // Пропуск при ошибке
                    }
                }
            }

            return elements;
        }

        // Загрузка типов элементов
        private void LoadTypes()
        {
            _typeNodes.Clear();

            // Получение типов из предварительно выбранных элементов или из связанных моделей
            if (_preSelectedElements != null && _preSelectedElements.Any())
            {
                LoadTypesFromPreSelectedElements();
            }
            else
            {
                LoadTypesFromLinkedModels();
            }

            // Обновление состояния родительских узлов
            foreach (var categoryNode in _typeNodes)
            {
                UpdateCategoryNodeState(categoryNode);
            }
        }

        // Загрузка типов из предварительно выбранных элементов
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

            // Создание узлов дерева
            foreach (var category in typesByCategory.OrderBy(c => c.Key))
            {
                var categoryNode = new TreeNode(category.Key) { IsExpanded = true };
                _typeNodes.Add(categoryNode);

                foreach (var type in category.Value.OrderBy(t => t.Key))
                {
                    var typeNode = categoryNode.AddChild(type.Key, type.Value);

                    // Восстановление предыдущего выбора
                    if (_filterSettings.SelectedTypeIds.Any(id => id.Value == type.Value.Value))
                    {
                        typeNode.SetCheckedSilent(true);
                    }
                }
            }
        }

        // Загрузка типов из связанных моделей
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

                    // Проверка, существует ли уже категория в дереве
                    var categoryNode = _typeNodes.FirstOrDefault(n => n.Name == categoryName);
                    if (categoryNode == null)
                    {
                        categoryNode = new TreeNode(categoryName) { IsExpanded = true };
                        _typeNodes.Add(categoryNode);
                    }

                    // Получение типов для категории
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
                            // Проверка, существует ли уже тип
                            if (!categoryNode.Children.Any(c => c.Name == type.Name))
                            {
                                var typeNode = categoryNode.AddChild(type.Name, type.Id);

                                // Восстановление предыдущего выбора
                                if (_filterSettings.SelectedTypeIds.Any(id => id.Value == type.Id.Value))
                                {
                                    typeNode.SetCheckedSilent(true);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Пропуск при ошибке получения типов
                    }
                }
            }
        }

        // Загрузка параметров для выбранных типов
        private void LoadParameters()
        {
            _parameterNodes.Clear();

            // Получение выбранных типов
            var selectedTypeIds = GetSelectedTypeIds();

            // Получение элементов для анализа
            var elements = GetElementsToFilter();

            // Фильтрация по выбранным типам
            if (selectedTypeIds.Any())
            {
                var typeIdValues = selectedTypeIds.Select(id => id.Value).ToHashSet();
                elements = elements.Where(e => typeIdValues.Contains(e.GetTypeId().Value)).ToList();
            }

            // Сбор параметров и их значений
            var parameterValues = new Dictionary<string, HashSet<string>>();

            foreach (var element in elements.Take(1000))
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param.Definition == null || param.IsReadOnly) continue;
                    if (param.StorageType == StorageType.None) continue;

                    var paramName = param.Definition.Name;
                    var value = param.GetStringValue();

                    if (string.IsNullOrWhiteSpace(value)) continue;

                    if (!parameterValues.ContainsKey(paramName))
                    {
                        parameterValues[paramName] = new HashSet<string>();
                    }

                    parameterValues[paramName].Add(value);
                }
            }

            // Создание узлов дерева параметров
            foreach (var param in parameterValues.OrderBy(p => p.Key))
            {
                var paramNode = new TreeNode(param.Key) { IsExpanded = false };

                foreach (var value in param.Value.OrderBy(v => v))
                {
                    var valueNode = paramNode.AddChild(value, value);

                    // Восстановление предыдущего выбора
                    if (_filterSettings.SelectedParameterValues.TryGetValue(param.Key, out var selectedValues))
                    {
                        if (selectedValues.Contains(value))
                        {
                            valueNode.SetCheckedSilent(true);
                        }
                    }
                }

                // Обновление состояния родительского узла
                UpdateCategoryNodeState(paramNode);

                _parameterNodes.Add(paramNode);
            }
        }

        // Получение ID выбранных типов
        private List<ElementId> GetSelectedTypeIds()
        {
            return _typeNodes
                .SelectMany(c => c.Children)
                .Where(t => t.IsChecked == true && t.Tag is ElementId)
                .Select(t => (ElementId)t.Tag)
                .ToList();
        }

        // Получение имён выбранных типов
        private List<string> GetSelectedTypeNames()
        {
            return _typeNodes
                .SelectMany(c => c.Children)
                .Where(t => t.IsChecked == true)
                .Select(t => t.Name)
                .ToList();
        }

        // Получение выбранных значений параметров
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

        // Подсчёт отфильтрованных элементов
        private int CountFilteredElements()
        {
            var selectedTypeNames = GetSelectedTypeNames();
            var selectedParamValues = GetSelectedParameterValues();

            var elements = GetElementsToFilter();

            // Фильтрация по именам типов
            if (selectedTypeNames.Any())
            {
                elements = elements
                    .Where(e =>
                    {
                        var typeId = e.GetTypeId();
                        if (typeId == ElementId.InvalidElementId) return false;
                        var typeElement = e.Document.GetElement(typeId);
                        return typeElement != null && selectedTypeNames.Contains(typeElement.Name);
                    })
                    .ToList();
            }

            // Фильтрация по параметрам
            if (selectedParamValues.Any())
            {
                elements = elements.Where(e =>
                {
                    foreach (var kvp in selectedParamValues)
                    {
                        var param = e.LookupParameter(kvp.Key);
                        if (param == null) continue;

                        var value = param.GetStringValue();
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

        // Обновление состояния узла категории на основе дочерних
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

        // Применение фильтра
        private void ApplyFilter()
        {
            _filterSettings.SelectedTypeIds = GetSelectedTypeIds();
            _filterSettings.SelectedTypeNames = GetSelectedTypeNames();
            _filterSettings.SelectedParameterValues = GetSelectedParameterValues();
            _filterSettings.ElementCount = CountFilteredElements();
            _filterSettings.IsFilterEnabled = true;

            if (_window != null)
            {
                _window.DialogResult = true;
                _window.Close();
            }
        }

        // Отмена
        private void Cancel()
        {
            if (_window != null)
            {
                _window.DialogResult = false;
                _window.Close();
            }
        }

        // Выбрать все типы
        private void SelectAllTypes()
        {
            foreach (var categoryNode in _typeNodes)
            {
                categoryNode.IsChecked = true;
            }
        }

        // Снять выбор со всех типов
        private void DeselectAllTypes()
        {
            foreach (var categoryNode in _typeNodes)
            {
                categoryNode.IsChecked = false;
            }
        }

        // Выбрать все параметры
        private void SelectAllParameters()
        {
            foreach (var paramNode in _parameterNodes)
            {
                paramNode.IsChecked = true;
            }
        }

        // Снять выбор со всех параметров
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
