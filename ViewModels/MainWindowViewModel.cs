using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OpeningTask.Helpers;
using OpeningTask.Models;
using OpeningTask.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace OpeningTask.ViewModels
{
    /// <summary>
    /// ViewModel for main plugin window
    /// </summary>
    public class MainWindowViewModel : BaseViewModel
    {
        private readonly UIDocument _uiDoc;
        private readonly Document _doc;
        private Window _window;

        // External event for Revit API operations
        private ExternalEvent _externalEvent;
        private CuboidPlacementEventHandler _eventHandler;

        // Model section
        private ObservableCollection<LinkedModelInfo> _mepLinkedModels;
        private ObservableCollection<LinkedModelInfo> _arKrLinkedModels;

        // Filter section - MEP systems
        private bool _isMepSelectedMode;
        private bool _isMepFilterMode;
        private FilterSettings _mepFilterSettings;
        private int _mepElementCount;

        // Filter section - walls
        private bool _isWallFilterEnabled;
        private FilterSettings _wallFilterSettings;
        private int _wallElementCount;

        // Filter section - floors
        private bool _isFloorFilterEnabled;
        private FilterSettings _floorFilterSettings;
        private int _floorElementCount;

        // Cuboid section
        private bool _roundDimensions;
        private int _roundDimensionsValue = 50;
        private bool _roundElevation;
        private int _roundElevationValue = 10;
        private bool _mergeIntersecting;

        // Cuboid types
        private bool _usePipeRound = true;
        private bool _usePipeRectangular;
        private bool _useDuctRound;
        private bool _useDuctRectangular;
        private bool _useTrayAll;

        private CuboidSettings _cuboidSettings;
        private int _minOffset = 30;
        private int _protrusion = 100;

        // Commands
        public ICommand OpenMepFilterCommand { get; }
        public ICommand OpenWallFilterCommand { get; }
        public ICommand OpenFloorFilterCommand { get; }
        public ICommand SelectMepElementsCommand { get; }
        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand OpenInstructionCommand { get; }

        public MainWindowViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc ?? throw new ArgumentNullException(nameof(uiDoc));
            _doc = uiDoc.Document;

            // Initialize collections
            _mepLinkedModels = new ObservableCollection<LinkedModelInfo>();
            _arKrLinkedModels = new ObservableCollection<LinkedModelInfo>();

            // Initialize filter settings
            _mepFilterSettings = new FilterSettings();
            _wallFilterSettings = new FilterSettings();
            _floorFilterSettings = new FilterSettings();

            // Initialize commands
            OpenMepFilterCommand = new RelayCommand(OpenMepFilter, CanOpenMepFilter);
            OpenWallFilterCommand = new RelayCommand(OpenWallFilter, () => IsWallFilterEnabled);
            OpenFloorFilterCommand = new RelayCommand(OpenFloorFilter, () => IsFloorFilterEnabled);
            SelectMepElementsCommand = new RelayCommand(SelectMepElements, CanSelectMepElements);
            OkCommand = new RelayCommand(ExecuteOk);
            CancelCommand = new RelayCommand(ExecuteCancel);
            OpenInstructionCommand = new RelayCommand(OpenInstruction);

            _cuboidSettings = new CuboidSettings();

            // Initialize external event handler
            _eventHandler = new CuboidPlacementEventHandler();
            _eventHandler.OperationCompleted += OnPlacementCompleted;
            _externalEvent = ExternalEvent.Create(_eventHandler);

            // Load linked models
            LoadLinkedModels();
        }

        public void SetWindow(Window window)
        {
            _window = window;
            if (_window != null)
            {
                _eventHandler.UiDispatcher = _window.Dispatcher;
            }
        }

        #region Model section properties

        public ObservableCollection<LinkedModelInfo> MepLinkedModels
        {
            get => _mepLinkedModels;
            set => SetProperty(ref _mepLinkedModels, value);
        }

        public ObservableCollection<LinkedModelInfo> ArKrLinkedModels
        {
            get => _arKrLinkedModels;
            set => SetProperty(ref _arKrLinkedModels, value);
        }

        #endregion

        #region Filter section properties - MEP systems

        public bool IsMepSelectedMode
        {
            get => _isMepSelectedMode;
            set
            {
                if (SetProperty(ref _isMepSelectedMode, value))
                {
                    // Update element count when mode changes
                    UpdateMepElementCount();
                }
            }
        }

        public bool IsMepFilterMode
        {
            get => _isMepFilterMode;
            set
            {
                if (SetProperty(ref _isMepFilterMode, value))
                {
                    _mepFilterSettings.IsFilterEnabled = value;
                    // Update element count when mode changes
                    UpdateMepElementCount();
                }
            }
        }

        public FilterSettings MepFilterSettings
        {
            get => _mepFilterSettings;
            set => SetProperty(ref _mepFilterSettings, value);
        }

        public int MepElementCount
        {
            get => _mepElementCount;
            set
            {
                if (SetProperty(ref _mepElementCount, value))
                {
                    OnPropertyChanged(nameof(MepElementCountText));
                }
            }
        }

        public string MepElementCountText => $"MEP Systems ({MepElementCount})";

        #endregion

        #region Filter section properties - walls

        public bool IsWallFilterEnabled
        {
            get => _isWallFilterEnabled;
            set
            {
                if (SetProperty(ref _isWallFilterEnabled, value))
                {
                    _wallFilterSettings.IsFilterEnabled = value;
                }
            }
        }

        public FilterSettings WallFilterSettings
        {
            get => _wallFilterSettings;
            set => SetProperty(ref _wallFilterSettings, value);
        }

        public int WallElementCount
        {
            get => _wallElementCount;
            set
            {
                if (SetProperty(ref _wallElementCount, value))
                {
                    OnPropertyChanged(nameof(WallElementCountText));
                }
            }
        }

        public string WallElementCountText => $"Walls ({WallElementCount})";

        #endregion

        #region Filter section properties - floors

        public bool IsFloorFilterEnabled
        {
            get => _isFloorFilterEnabled;
            set
            {
                if (SetProperty(ref _isFloorFilterEnabled, value))
                {
                    _floorFilterSettings.IsFilterEnabled = value;
                }
            }
        }

        public FilterSettings FloorFilterSettings
        {
            get => _floorFilterSettings;
            set => SetProperty(ref _floorFilterSettings, value);
        }

        public int FloorElementCount
        {
            get => _floorElementCount;
            set
            {
                if (SetProperty(ref _floorElementCount, value))
                {
                    OnPropertyChanged(nameof(FloorElementCountText));
                }
            }
        }

        public string FloorElementCountText => $"Floors ({FloorElementCount})";

        #endregion

        #region Cuboid section properties

        public bool RoundDimensions
        {
            get => _roundDimensions;
            set => SetProperty(ref _roundDimensions, value);
        }

        public int RoundDimensionsValue
        {
            get => _roundDimensionsValue;
            set => SetProperty(ref _roundDimensionsValue, value);
        }

        public bool RoundElevation
        {
            get => _roundElevation;
            set => SetProperty(ref _roundElevation, value);
        }

        public int RoundElevationValue
        {
            get => _roundElevationValue;
            set => SetProperty(ref _roundElevationValue, value);
        }

        public bool MergeIntersecting
        {
            get => _mergeIntersecting;
            set => SetProperty(ref _mergeIntersecting, value);
        }

        public bool UsePipeRound
        {
            get => _usePipeRound;
            set => SetProperty(ref _usePipeRound, value);
        }

        public bool UsePipeRectangular
        {
            get => _usePipeRectangular;
            set => SetProperty(ref _usePipeRectangular, value);
        }

        public bool UseDuctRound
        {
            get => _useDuctRound;
            set => SetProperty(ref _useDuctRound, value);
        }

        public bool UseDuctRectangular
        {
            get => _useDuctRectangular;
            set => SetProperty(ref _useDuctRectangular, value);
        }

        public bool UseTrayAll
        {
            get => _useTrayAll;
            set => SetProperty(ref _useTrayAll, value);
        }

        /// <summary>
        /// Настройки кубиков
        /// </summary>
        public CuboidSettings CuboidSettings
        {
            get => _cuboidSettings;
            set => SetProperty(ref _cuboidSettings, value);
        }

        /// <summary>
        /// Минимальный отступ от элемента (мм)
        /// </summary>
        public int MinOffset
        {
            get => _minOffset;
            set
            {
                if (SetProperty(ref _minOffset, value))
                {
                    _cuboidSettings.MinOffset = value;
                }
            }
        }

        /// <summary>
        /// Выступ от элемента вставки (мм)
        /// </summary>
        public int Protrusion
        {
            get => _protrusion;
            set
            {
                if (SetProperty(ref _protrusion, value))
                {
                    _cuboidSettings.Protrusion = value;
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Load linked models
        /// </summary>
        private void LoadLinkedModels()
        {
            var linkedModels = Functions.GetAllLinkedModels(_doc);

            _mepLinkedModels.Clear();
            _arKrLinkedModels.Clear();

            foreach (var model in linkedModels)
            {
                // Add to both collections (user chooses which model for what purpose)
                var mepModel = new LinkedModelInfo(model.LinkInstance);
                var arKrModel = new LinkedModelInfo(model.LinkInstance);

                _mepLinkedModels.Add(mepModel);
                _arKrLinkedModels.Add(arKrModel);
            }
        }

        /// <summary>
        /// Check if can open MEP filter
        /// </summary>
        private bool CanOpenMepFilter()
        {
            return IsMepFilterMode;
        }

        /// <summary>
        /// Check if can select MEP elements
        /// </summary>
        private bool CanSelectMepElements()
        {
            return IsMepSelectedMode && _mepLinkedModels.Any(m => m.IsSelected);
        }

        /// <summary>
        /// Update MEP element count based on current filter settings
        /// </summary>
        private void UpdateMepElementCount()
        {
            int count = 0;

            // Count from manually selected elements
            if (IsMepSelectedMode && _mepFilterSettings.SelectedLinkedElements.Any())
            {
                count = _mepFilterSettings.SelectedLinkedElements.Count;
            }

            // If filter is also enabled, the FilterWindow will refine this count
            if (IsMepFilterMode && _mepFilterSettings.ElementCount > 0)
            {
                // Use filter count if available
                count = _mepFilterSettings.ElementCount;
            }

            MepElementCount = count;
        }

        /// <summary>
        /// Open MEP filter window
        /// </summary>
        private void OpenMepFilter()
        {
            var selectedMepModels = _mepLinkedModels.Where(m => m.IsSelected).ToList();
            
            // If we have manually selected elements, use them as base for filtering
            var preSelectedElements = IsMepSelectedMode ? _mepFilterSettings.SelectedLinkedElements : null;

            if (!selectedMepModels.Any() && (preSelectedElements == null || !preSelectedElements.Any()))
            {
                MessageBox.Show("Сначала выберите модели инженерных систем или элементы на виде", 
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var filterViewModel = new FilterWindowViewModel(
                selectedMepModels, 
                Functions.MepCategories, 
                _mepFilterSettings,
                "Фильтр инженерных систем",
                preSelectedElements);

            var filterWindow = new Views.FilterWindow(filterViewModel);
            filterWindow.Owner = _window;
            
            if (filterWindow.ShowDialog() == true)
            {
                _mepFilterSettings = filterViewModel.FilterSettings;
                MepElementCount = _mepFilterSettings.ElementCount;
            }
        }

        /// <summary>
        /// Open wall filter window
        /// </summary>
        private void OpenWallFilter()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedArKrModels.Any())
            {
                MessageBox.Show("Select at least one AP / KP model first", 
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var filterViewModel = new FilterWindowViewModel(
                selectedArKrModels, 
                new[] { BuiltInCategory.OST_Walls }, 
                _wallFilterSettings,
                "Wall Filter");

            var filterWindow = new Views.FilterWindow(filterViewModel);
            filterWindow.Owner = _window;
            
            if (filterWindow.ShowDialog() == true)
            {
                _wallFilterSettings = filterViewModel.FilterSettings;
                WallElementCount = _wallFilterSettings.ElementCount;
            }
        }

        /// <summary>
        /// Open floor filter window
        /// </summary>
        private void OpenFloorFilter()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedArKrModels.Any())
            {
                MessageBox.Show("Select at least one AP / KP model first", 
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var filterViewModel = new FilterWindowViewModel(
                selectedArKrModels, 
                new[] { BuiltInCategory.OST_Floors }, 
                _floorFilterSettings,
                "Floor Filter");

            var filterWindow = new Views.FilterWindow(filterViewModel);
            filterWindow.Owner = _window;
            
            if (filterWindow.ShowDialog() == true)
            {
                _floorFilterSettings = filterViewModel.FilterSettings;
                FloorElementCount = _floorFilterSettings.ElementCount;
            }
        }

        /// <summary>
        /// Select MEP elements from linked models
        /// </summary>
        private void SelectMepElements()
        {
            // Check if any MEP models are selected
            var selectedMepModels = _mepLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedMepModels.Any())
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель инженерных систем", 
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Hide window for selection
                _window?.Hide();

                // Select elements from linked models
                var selectedElements = Functions.SelectElementsFromLinkedModels(
                    _uiDoc, 
                    selectedMepModels, 
                    Functions.MepCategories);

                // Filter by MEP categories
                var filteredElements = selectedElements
                    .Where(e => e.LinkedElement != null && 
                                e.LinkedElement.Category != null &&
                                Functions.MepCategories.Any(c => 
                                    e.LinkedElement.Category.Id.Value == (long)c))
                    .ToList();

                _mepFilterSettings.SelectedLinkedElements = filteredElements;
                _mepFilterSettings.FilterType = FilterType.Selected;
                MepElementCount = filteredElements.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting elements: {ex.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _window?.Show();
            }
        }

        /// <summary>
        /// Execute OK command
        /// </summary>
        private void ExecuteOk()
        {
            RevitTrace.Info("UI OK: clicked");
            // Валидация
            if (!_mepLinkedModels.Any(m => m.IsSelected))
            {
                MessageBox.Show("Выберите хотя бы одну модель инженерных систем",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_arKrLinkedModels.Any(m => m.IsSelected))
            {
                MessageBox.Show("Выберите хотя бы одну модель АР/КР",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Собираем элементы для проверки пересечений
            RevitTrace.Info("UI OK: collecting elements");
            var mepElements = CollectMepElements();
            var wallElements = CollectWallElements();
            var floorElements = CollectFloorElements();

            RevitTrace.Info($"UI OK: collected mep={mepElements.Count}, walls={wallElements.Count}, floors={floorElements.Count}");

            if (!mepElements.Any())
            {
                MessageBox.Show("Не выбраны элементы инженерных систем",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!wallElements.Any() && !floorElements.Any())
            {
                MessageBox.Show("Не выбраны стены или перекрытия для поиска пересечений",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Настройка параметров кубиков
            _cuboidSettings.RoundingValue = RoundDimensionsValue;
            _cuboidSettings.MinOffset = MinOffset;
            _cuboidSettings.Protrusion = Protrusion;
            _cuboidSettings.UsePipeRoundCuboid = UsePipeRound;
            _cuboidSettings.UseDuctRoundCuboid = UseDuctRound;

            // Создаём запрос для внешнего события
            RevitTrace.Info("UI OK: building request");
            _eventHandler.Request = new CuboidPlacementRequest
            {
                MepElements = mepElements,
                WallElements = wallElements,
                FloorElements = floorElements,
                Settings = _cuboidSettings
            };

            // Запускаем внешнее событие
            RevitTrace.Info("UI OK: raising ExternalEvent");
            _externalEvent.Raise();
            RevitTrace.Info("UI OK: ExternalEvent.Raise returned");

            // Закрываем окно (результат покажется после выполнения)
            if (_window != null)
            {
                try
                {
                    RevitTrace.Info("UI OK: closing MainWindow");
                    var canSetDialogResult = false;
                    try
                    {
                        // DialogResult доступен только если окно показано как диалоговое
                        // (иначе выбрасывается InvalidOperationException)
                        canSetDialogResult = _window.IsLoaded;
                    }
                    catch
                    {
                        canSetDialogResult = false;
                    }

                    if (canSetDialogResult)
                    {
                        try
                        {
                            _window.DialogResult = true;
                        }
                        catch (InvalidOperationException)
                        {
                            // Окно не диалоговое в текущем контексте
                        }
                    }

                    _window.Close();
                    RevitTrace.Info("UI OK: MainWindow closed");
                }
                catch (Exception ex)
                {
                    RevitTrace.Error("UI OK: exception during MainWindow close", ex);
                    throw;
                }
            }
        }

        /// <summary>
        /// Обработчик завершения операции размещения кубиков
        /// </summary>
        private void OnPlacementCompleted(bool success, int count, string errorMessage)
        {
            if (success)
            {
                if (count > 0)
                {
                    MessageBox.Show($"Создано кубиков: {count}",
                        "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else if (!string.IsNullOrEmpty(errorMessage))
                {
                    MessageBox.Show(errorMessage,
                        "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else
            {
                MessageBox.Show($"Ошибка: {errorMessage}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Собрать MEP элементы для проверки
        /// </summary>
        private List<LinkedElementInfo> CollectMepElements()
        {
            var result = new List<LinkedElementInfo>();

            if (IsMepSelectedMode && _mepFilterSettings.SelectedLinkedElements.Any())
            {
                result.AddRange(_mepFilterSettings.SelectedLinkedElements);
            }

            if (IsMepFilterMode && _mepFilterSettings.IsFilterEnabled)
            {
                foreach (var linkedModel in _mepLinkedModels.Where(m => m.IsSelected && m.IsLoaded))
                {
                    var elements = Functions.FilterElements(
                        linkedModel.LinkedDocument,
                        _mepFilterSettings,
                        Functions.MepCategories);

                    foreach (var element in elements)
                    {
                        result.Add(new LinkedElementInfo
                        {
                            LinkInstance = linkedModel.LinkInstance,
                            LinkedElementId = element.Id,
                            LinkedElement = element,
                            LinkedDocument = linkedModel.LinkedDocument
                        });
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Собрать стены для проверки
        /// </summary>
        private List<LinkedElementInfo> CollectWallElements()
        {
            var result = new List<LinkedElementInfo>();

            if (!IsWallFilterEnabled) return result;

            foreach (var linkedModel in _arKrLinkedModels.Where(m => m.IsSelected && m.IsLoaded))
            {
                var elements = Functions.FilterElements(
                    linkedModel.LinkedDocument,
                    _wallFilterSettings,
                    new[] { BuiltInCategory.OST_Walls });

                foreach (var element in elements)
                {
                    result.Add(new LinkedElementInfo
                    {
                        LinkInstance = linkedModel.LinkInstance,
                        LinkedElementId = element.Id,
                        LinkedElement = element,
                        LinkedDocument = linkedModel.LinkedDocument
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Собрать перекрытия для проверки
        /// </summary>
        private List<LinkedElementInfo> CollectFloorElements()
        {
            var result = new List<LinkedElementInfo>();

            if (!IsFloorFilterEnabled) return result;

            foreach (var linkedModel in _arKrLinkedModels.Where(m => m.IsSelected && m.IsLoaded))
            {
                var elements = Functions.FilterElements(
                    linkedModel.LinkedDocument,
                    _floorFilterSettings,
                    new[] { BuiltInCategory.OST_Floors });

                foreach (var element in elements)
                {
                    result.Add(new LinkedElementInfo
                    {
                        LinkInstance = linkedModel.LinkInstance,
                        LinkedElementId = element.Id,
                        LinkedElement = element,
                        LinkedDocument = linkedModel.LinkedDocument
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Execute Cancel command
        /// </summary>
        private void ExecuteCancel()
        {
            if (_window != null)
            {
                _window.DialogResult = false;
                _window.Close();
            }
        }

        /// <summary>
        /// Open instruction
        /// </summary>
        private void OpenInstruction()
        {
            try
            {
                System.Diagnostics.Process.Start("https://example.com/instruction");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open instruction: {ex.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion
    }
}
