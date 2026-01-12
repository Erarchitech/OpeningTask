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
    // ViewModel для главного окна плагина
    public class MainWindowViewModel : BaseViewModel
    {
        private readonly UIDocument _uiDoc;
        private readonly Document _doc;
        private Window _window;

        // Внешнее событие для операций Revit API
        private ExternalEvent _externalEvent;
        private CuboidPlacementEventHandler _eventHandler;

        // Секция моделей
        private ObservableCollection<LinkedModelInfo> _mepLinkedModels;
        private ObservableCollection<LinkedModelInfo> _arKrLinkedModels;

        // Секция фильтра - MEP системы
        private bool _isMepSelectedMode;
        private bool _isMepFilterMode;
        private FilterSettings _mepFilterSettings;
        private int _mepElementCount;

        // Секция фильтра - стены
        private bool _isWallSelectedMode;
        private bool _isWallFilterMode;
        private FilterSettings _wallFilterSettings;
        private int _wallElementCount;

        // Секция фильтра - перекрытия
        private bool _isFloorSelectedMode;
        private bool _isFloorFilterMode;
        private FilterSettings _floorFilterSettings;
        private int _floorElementCount;

        // Секция боксов
        private bool _roundDimensions;
        private int _roundDimensionsValue = 50;
        private bool _roundElevation;
        private int _roundElevationValue = 10;
        private bool _mergeIntersecting;

        // Типы боксов
        private bool _usePipeRound = true;
        private bool _usePipeRectangular;
        private bool _useDuctRound;
        private bool _useDuctRectangular;
        private bool _useTrayAll;

        private CuboidSettings _cuboidSettings;
        private int _minOffset = 30;
        private int _protrusion = 100;

        // Команды
        public ICommand OpenMepFilterCommand { get; }
        public ICommand OpenWallFilterCommand { get; }
        public ICommand OpenFloorFilterCommand { get; }
        public ICommand SelectMepElementsCommand { get; }
        public ICommand SelectWallElementsCommand { get; }
        public ICommand SelectFloorElementsCommand { get; }
        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand OpenInstructionCommand { get; }

        public MainWindowViewModel(UIDocument uiDoc)
        {
            _uiDoc = uiDoc ?? throw new ArgumentNullException(nameof(uiDoc));
            _doc = uiDoc.Document;

            // Инициализация коллекций
            _mepLinkedModels = new ObservableCollection<LinkedModelInfo>();
            _arKrLinkedModels = new ObservableCollection<LinkedModelInfo>();

            // Инициализация настроек фильтра
            _mepFilterSettings = new FilterSettings();
            _wallFilterSettings = new FilterSettings();
            _floorFilterSettings = new FilterSettings();

            // Инициализация команд
            OpenMepFilterCommand = new RelayCommand(OpenMepFilter, CanOpenMepFilter);
            OpenWallFilterCommand = new RelayCommand(OpenWallFilter, CanOpenWallFilter);
            OpenFloorFilterCommand = new RelayCommand(OpenFloorFilter, CanOpenFloorFilter);
            SelectMepElementsCommand = new RelayCommand(SelectMepElements, CanSelectMepElements);
            SelectWallElementsCommand = new RelayCommand(SelectWallElements, CanSelectWallElements);
            SelectFloorElementsCommand = new RelayCommand(SelectFloorElements, CanSelectFloorElements);
            OkCommand = new RelayCommand(ExecuteOk);
            CancelCommand = new RelayCommand(ExecuteCancel);
            OpenInstructionCommand = new RelayCommand(OpenInstruction);

            _cuboidSettings = new CuboidSettings();

            // Инициализация обработчика внешнего события
            _eventHandler = new CuboidPlacementEventHandler();
            _eventHandler.OperationCompleted += OnPlacementCompleted;
            _externalEvent = ExternalEvent.Create(_eventHandler);

            // Загрузка связанных моделей
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

        #region Свойства секции моделей

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

        #region Свойства секции фильтра - MEP системы

        public bool IsMepSelectedMode
        {
            get => _isMepSelectedMode;
            set
            {
                if (SetProperty(ref _isMepSelectedMode, value))
                {
                    // Обновление счётчика элементов при смене режима
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
                    // Обновление счётчика элементов при смене режима
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

        public string MepElementCountText => $"Элементы ИС ({MepElementCount})";

        #endregion

        #region Свойства секции фильтра - стены

        public bool IsWallSelectedMode
        {
            get => _isWallSelectedMode;
            set
            {
                if (SetProperty(ref _isWallSelectedMode, value))
                {
                    UpdateWallElementCount();
                }
            }
        }

        public bool IsWallFilterMode
        {
            get => _isWallFilterMode;
            set
            {
                if (SetProperty(ref _isWallFilterMode, value))
                {
                    _wallFilterSettings.IsFilterEnabled = value;
                    UpdateWallElementCount();
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

        public string WallElementCountText => $"Стены ({WallElementCount})";

        #endregion

        #region Свойства секции фильтра - перекрытия

        public bool IsFloorSelectedMode
        {
            get => _isFloorSelectedMode;
            set
            {
                if (SetProperty(ref _isFloorSelectedMode, value))
                {
                    UpdateFloorElementCount();
                }
            }
        }

        public bool IsFloorFilterMode
        {
            get => _isFloorFilterMode;
            set
            {
                if (SetProperty(ref _isFloorFilterMode, value))
                {
                    _floorFilterSettings.IsFilterEnabled = value;
                    UpdateFloorElementCount();
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

        public string FloorElementCountText => $"Перекрытия ({FloorElementCount})";

        #endregion

        #region Свойства секции боксов

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
            set
            {
                if (SetProperty(ref _usePipeRound, value))
                {
                    // Взаимоисключение
                    if (value)
                        UsePipeRectangular = false;

                    // Синхронизация с настройками
                    _cuboidSettings.UsePipeRoundCuboid = value;
                }
            }
        }

        public bool UsePipeRectangular
        {
            get => _usePipeRectangular;
            set
            {
                if (SetProperty(ref _usePipeRectangular, value))
                {
                    // Взаимоисключение
                    if (value)
                        UsePipeRound = false;

                    // Прямоугольный = не круглый
                    _cuboidSettings.UsePipeRoundCuboid = !value;
                }
            }
        }

        public bool UseDuctRound
        {
            get => _useDuctRound;
            set
            {
                if (SetProperty(ref _useDuctRound, value))
                {
                    // Взаимоисключение
                    if (value)
                        UseDuctRectangular = false;

                    // Синхронизация с настройками
                    _cuboidSettings.UseDuctRoundCuboid = value;
                }
            }
        }

        public bool UseDuctRectangular
        {
            get => _useDuctRectangular;
            set
            {
                if (SetProperty(ref _useDuctRectangular, value))
                {
                    // Взаимоисключение
                    if (value)
                        UseDuctRound = false;

                    // Прямоугольный = не круглый
                    _cuboidSettings.UseDuctRoundCuboid = !value;
                }
            }
        }

        public bool UseTrayAll
        {
            get => _useTrayAll;
            set => SetProperty(ref _useTrayAll, value);
        }

        // Настройки боксов
        public CuboidSettings CuboidSettings
        {
            get => _cuboidSettings;
            set => SetProperty(ref _cuboidSettings, value);
        }

        // Минимальный отступ от элемента (мм)
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

        // Выступ от элемента вставки (мм)
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

        #region Методы

        // Загрузка связанных моделей
        private void LoadLinkedModels()
        {
            var linkedModels = Functions.GetAllLinkedModels(_doc);

            _mepLinkedModels.Clear();
            _arKrLinkedModels.Clear();

            foreach (var model in linkedModels)
            {
                // Добавление в обе коллекции
                var mepModel = new LinkedModelInfo(model.LinkInstance);
                var arKrModel = new LinkedModelInfo(model.LinkInstance);

                _mepLinkedModels.Add(mepModel);
                _arKrLinkedModels.Add(arKrModel);
            }
        }

        // Проверка доступности открытия фильтра MEP
        private bool CanOpenMepFilter()
        {
            return IsMepFilterMode;
        }

        private bool CanOpenWallFilter()
        {
            return IsWallFilterMode;
        }

        private bool CanOpenFloorFilter()
        {
            return IsFloorFilterMode;
        }

        // Проверка доступности выбора MEP элементов
        private bool CanSelectMepElements()
        {
            return IsMepSelectedMode && _mepLinkedModels.Any(m => m.IsSelected);
        }

        private bool CanSelectWallElements()
        {
            return IsWallSelectedMode && _arKrLinkedModels.Any(m => m.IsSelected);
        }

        private bool CanSelectFloorElements()
        {
            return IsFloorSelectedMode && _arKrLinkedModels.Any(m => m.IsSelected);
        }

        // Обновление счётчика MEP элементов
        private void UpdateMepElementCount()
        {
            int count = 0;

            // Подсчёт вручную выбранных элементов
            if (IsMepSelectedMode && _mepFilterSettings.SelectedLinkedElements.Any())
            {
                count = _mepFilterSettings.SelectedLinkedElements.Count;
            }

            // Если фильтр включён, используем счётчик фильтра
            if (IsMepFilterMode && _mepFilterSettings.ElementCount > 0)
            {
                count = _mepFilterSettings.ElementCount;
            }

            MepElementCount = count;
        }

        private void UpdateWallElementCount()
        {
            int count = 0;

            if (IsWallSelectedMode && _wallFilterSettings.SelectedLinkedElements.Any())
            {
                count = _wallFilterSettings.SelectedLinkedElements.Count;
            }

            if (IsWallFilterMode && _wallFilterSettings.ElementCount > 0)
            {
                count = _wallFilterSettings.ElementCount;
            }

            WallElementCount = count;
        }

        private void UpdateFloorElementCount()
        {
            int count = 0;

            if (IsFloorSelectedMode && _floorFilterSettings.SelectedLinkedElements.Any())
            {
                count = _floorFilterSettings.SelectedLinkedElements.Count;
            }

            if (IsFloorFilterMode && _floorFilterSettings.ElementCount > 0)
            {
                count = _floorFilterSettings.ElementCount;
            }

            FloorElementCount = count;
        }

        // Открытие окна фильтра MEP
        private void OpenMepFilter()
        {
            var selectedMepModels = _mepLinkedModels.Where(m => m.IsSelected).ToList();
            
            // Если есть вручную выбранные элементы, используем их как базу
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

        // Открытие окна фильтра стен
        private void OpenWallFilter()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            var preSelectedElements = IsWallSelectedMode ? _wallFilterSettings.SelectedLinkedElements : null;

            if (!selectedArKrModels.Any() && (preSelectedElements == null || !preSelectedElements.Any()))
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель АР/КР или выбранные стены на виде",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var filterViewModel = new FilterWindowViewModel(
                selectedArKrModels, 
                new[] { BuiltInCategory.OST_Walls }, 
                _wallFilterSettings,
                "Фильтр стен",
                preSelectedElements);

            var filterWindow = new Views.FilterWindow(filterViewModel);
            filterWindow.Owner = _window;
            
            if (filterWindow.ShowDialog() == true)
            {
                _wallFilterSettings = filterViewModel.FilterSettings;
                WallElementCount = _wallFilterSettings.ElementCount;
            }
        }

        // Открытие окна фильтра перекрытий
        private void OpenFloorFilter()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            var preSelectedElements = IsFloorSelectedMode ? _floorFilterSettings.SelectedLinkedElements : null;

            if (!selectedArKrModels.Any() && (preSelectedElements == null || !preSelectedElements.Any()))
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель АР/КР или выбранные перекрытия на виде",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var filterViewModel = new FilterWindowViewModel(
                selectedArKrModels, 
                new[] { BuiltInCategory.OST_Floors }, 
                _floorFilterSettings,
                "Фильтр перекрытий",
                preSelectedElements);

            var filterWindow = new Views.FilterWindow(filterViewModel);
            filterWindow.Owner = _window;
            
            if (filterWindow.ShowDialog() == true)
            {
                _floorFilterSettings = filterViewModel.FilterSettings;
                FloorElementCount = _floorFilterSettings.ElementCount;
            }
        }

        private void SelectWallElements()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedArKrModels.Any())
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель АР/КР",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _window?.Hide();

                var selectedElements = Functions.SelectElementsFromLinkedModels(
                    _uiDoc,
                    selectedArKrModels,
                    new[] { BuiltInCategory.OST_Walls });

                var filteredElements = selectedElements
                    .Where(e => e.LinkedElement != null &&
                                e.LinkedElement.Category != null &&
                                e.LinkedElement.Category.Id.Value == (long)BuiltInCategory.OST_Walls)
                    .ToList();

                _wallFilterSettings.SelectedLinkedElements = filteredElements;
                _wallFilterSettings.FilterType = FilterType.Selected;
                WallElementCount = filteredElements.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка выбора элементов: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _window?.Show();
            }
        }

        private void SelectFloorElements()
        {
            var selectedArKrModels = _arKrLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedArKrModels.Any())
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель АР/КР",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _window?.Hide();

                var selectedElements = Functions.SelectElementsFromLinkedModels(
                    _uiDoc,
                    selectedArKrModels,
                    new[] { BuiltInCategory.OST_Floors });

                var filteredElements = selectedElements
                    .Where(e => e.LinkedElement != null &&
                                e.LinkedElement.Category != null &&
                                e.LinkedElement.Category.Id.Value == (long)BuiltInCategory.OST_Floors)
                    .ToList();

                _floorFilterSettings.SelectedLinkedElements = filteredElements;
                _floorFilterSettings.FilterType = FilterType.Selected;
                FloorElementCount = filteredElements.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка выбора элементов: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _window?.Show();
            }
        }

        // Выбор MEP элементов из связанных моделей
        private void SelectMepElements()
        {
            // Проверка, выбраны ли MEP модели
            var selectedMepModels = _mepLinkedModels.Where(m => m.IsSelected).ToList();
            if (!selectedMepModels.Any())
            {
                MessageBox.Show("Сначала выберите хотя бы одну модель инженерных систем", 
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Скрытие окна для выбора
                _window?.Hide();

                // Выбор элементов из связанных моделей
                var selectedElements = Functions.SelectElementsFromLinkedModels(
                    _uiDoc, 
                    selectedMepModels, 
                    Functions.MepCategories);

                // Фильтрация по MEP категориям
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
                MessageBox.Show($"Ошибка выбора элементов: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _window?.Show();
            }
        }

        // Выполнение команды OK
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

        // Обработчик завершения операции размещения боксов
        private void OnPlacementCompleted(bool success, int count, string errorMessage)
        {
            if (success)
            {
                try
                {
                    var duplicateIds = _eventHandler?.DuplicateCuboidIds;
                    if (duplicateIds != null && duplicateIds.Any())
                    {
                        RevitTrace.Info($"UI: showing duplicate cuboids warning, count={duplicateIds.Count}");

                        var win = new Views.DuplicateCuboidsReportWindow(duplicateIds.Select(x => (long)x.Value));

                        // MainWindow is closed right after ExternalEvent.Raise, so Owner can be null/invalid.
                        if (_window != null && _window.IsLoaded)
                        {
                            win.Owner = _window;
                        }
                        else
                        {
                            win.Topmost = true;
                        }

                        win.ShowDialog();
                    }
                }
                catch
                {
                    // ignore UI errors
                }

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

        // Сбор MEP элементов для проверки
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

        // Сбор стен для проверки
        private List<LinkedElementInfo> CollectWallElements()
        {
            return CollectHostElements(
                IsWallSelectedMode, IsWallFilterMode,
                _wallFilterSettings, BuiltInCategory.OST_Walls);
        }

        // Сбор перекрытий для проверки
        private List<LinkedElementInfo> CollectFloorElements()
        {
            return CollectHostElements(
                IsFloorSelectedMode, IsFloorFilterMode,
                _floorFilterSettings, BuiltInCategory.OST_Floors);
        }

        // Общий метод сбора элементов хоста (стены/перекрытия)
        private List<LinkedElementInfo> CollectHostElements(
            bool isSelectedMode, bool isFilterMode,
            FilterSettings filterSettings, BuiltInCategory category)
        {
            var result = new List<LinkedElementInfo>();
            var seen = new HashSet<string>();

            if (isSelectedMode && filterSettings.SelectedLinkedElements.Any())
            {
                foreach (var info in filterSettings.SelectedLinkedElements)
                {
                    if (info?.LinkInstance == null || info.LinkedElementId == null) continue;
                    var key = info.LinkInstance.Id.Value + ":" + info.LinkedElementId.Value;
                    if (seen.Add(key))
                        result.Add(info);
                }
            }

            if (isFilterMode && filterSettings.IsFilterEnabled)
            {
                foreach (var linkedModel in _arKrLinkedModels.Where(m => m.IsSelected && m.IsLoaded))
                {
                    var elements = Functions.FilterElements(
                        linkedModel.LinkedDocument,
                        filterSettings,
                        new[] { category });

                    foreach (var element in elements)
                    {
                        var key = linkedModel.LinkInstance.Id.Value + ":" + element.Id.Value;
                        if (seen.Add(key))
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
            }

            return result;
        }

        // Выполнение команды Отмена
        private void ExecuteCancel()
        {
            if (_window != null)
            {
                _window.DialogResult = false;
                _window.Close();
            }
        }

        // Открытие инструкции
        private void OpenInstruction()
        {
            try
            {
                System.Diagnostics.Process.Start("https://task.com/instruction");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть инструкцию: {ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion
    }
}
