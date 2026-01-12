# OpeningTask

Плагин для Autodesk Revit 2024, автоматизирующий создание заданий на отверстия в местах пересечения инженерных систем (MEP) со стенами и перекрытиями.

## Возможности

- Автоматический поиск пересечений MEP элементов (трубы, воздуховоды, кабельные лотки) со стенами и перекрытиями
- Размещение параметрических боксов (семейств) в точках пересечений
- Поддержка работы со связанными моделями (Revit Links)
- Гибкая фильтрация элементов по типам и параметрам
- Настройка размеров боксов с округлением и отступами
- Обнаружение дубликатов при повторном запуске

## Требования

- Autodesk Revit 2024
- .NET Framework 4.8
- Семейства боксов в папке `C:\OpeningTaskResources`

## Установка

1. Скомпилируйте проект
2. Скопируйте `OpeningTask.dll` и `OpeningTask.addin` в папку `%AppData%\Autodesk\Revit\Addins\2024`
3. Разместите файлы семейств (`.rfa`) в папке `C:\OpeningTaskResources`

## Структура проекта

```
OpeningTask/
??? Commands/
?   ??? OpeningTaskCommand.cs      # Точка входа плагина (IExternalCommand)
??? Helpers/
?   ??? CuboidPlacementEventHandler.cs  # Обработчик ExternalEvent для Revit API
?   ??? Extensions.cs              # Методы расширения (Parameter.GetStringValue)
?   ??? Functions.cs               # Вспомогательные функции (фильтрация, выбор элементов)
?   ??? RelayCommand.cs            # Реализация ICommand для MVVM
?   ??? RevitTrace.cs              # Логирование для отладки
??? Models/
?   ??? CuboidSettings.cs          # Настройки размещения боксов (отступы, округление)
?   ??? FilterSettings.cs          # Настройки фильтрации элементов
?   ??? IntersectionInfo.cs        # Информация о пересечении MEP с хостом
?   ??? LinkedModelInfo.cs         # Информация о связанной модели Revit
?   ??? TreeNode.cs                # Узел дерева для UI фильтрации
??? Services/
?   ??? CuboidPlacementService.cs  # Размещение боксов в точках пересечений
?   ??? IntersectionService.cs     # Поиск пересечений через Boolean операции
??? ViewModels/
?   ??? BaseViewModel.cs           # Базовый класс с INotifyPropertyChanged
?   ??? FilterWindowViewModel.cs   # ViewModel окна фильтрации
?   ??? MainWindowViewModel.cs     # ViewModel главного окна
??? Views/
?   ??? DuplicateCuboidsReportWindow.xaml  # Отчёт о дубликатах
?   ??? FilterWindow.xaml          # Окно фильтрации по типам/параметрам
?   ??? MainWindow.xaml            # Главное окно плагина
??? Resources/
?   ??? Images/                    # Иконки для UI
?   ??? Styles.xaml                # WPF стили
??? Interface.cs                   # IExternalApplication (создание кнопки на ленте)
??? OpeningTask.addin              # Манифест плагина для Revit
??? OpeningTask.csproj             # Файл проекта
```

## Описание компонентов

### Commands
| Файл | Описание |
|------|----------|
| `OpeningTaskCommand.cs` | Основная команда плагина. Реализует `IExternalCommand`, открывает главное окно. |

### Helpers
| Файл | Описание |
|------|----------|
| `CuboidPlacementEventHandler.cs` | Обработчик `IExternalEventHandler` для выполнения операций Revit API из UI потока. |
| `Extensions.cs` | Метод расширения `GetStringValue()` для получения строкового значения параметра. |
| `Functions.cs` | Статические методы: загрузка связей, фильтрация элементов, интерактивный выбор. |
| `RelayCommand.cs` | Реализация `ICommand` для привязки команд в MVVM. |
| `RevitTrace.cs` | Логирование в файл для отладки работы плагина. |

### Models
| Файл | Описание |
|------|----------|
| `CuboidSettings.cs` | Настройки боксов: округление размеров, отступы, выступы, GUID параметров. |
| `FilterSettings.cs` | Хранит выбранные типы, параметры и список отфильтрованных элементов. |
| `IntersectionInfo.cs` | Данные о пересечении: MEP элемент, хост, точка вставки, размеры. |
| `LinkedModelInfo.cs` | Обёртка над `RevitLinkInstance` с флагом выбора для UI. |
| `TreeNode.cs` | Узел дерева с поддержкой трёхсостояного чекбокса для иерархического выбора. |

### Services
| Файл | Описание |
|------|----------|
| `CuboidPlacementService.cs` | Размещает семейства боксов, задаёт параметры, ориентирует по направлению MEP. |
| `IntersectionService.cs` | Находит пересечения через `BooleanOperationsUtils`, вычисляет точку вставки. |

### ViewModels
| Файл | Описание |
|------|----------|
| `BaseViewModel.cs` | Базовый класс с `INotifyPropertyChanged` и методом `SetProperty`. |
| `FilterWindowViewModel.cs` | Логика окна фильтрации: загрузка типов, параметров, подсчёт элементов. |
| `MainWindowViewModel.cs` | Основная логика: выбор моделей, сбор элементов, запуск размещения. |

### Views
| Файл | Описание |
|------|----------|
| `MainWindow.xaml` | Главное окно с выбором моделей, настройками боксов, кнопками управления. |
| `FilterWindow.xaml` | Окно фильтрации с деревом типов и параметров. |
| `DuplicateCuboidsReportWindow.xaml` | Отчёт о найденных дубликатах боксов. |

### Корневые файлы
| Файл | Описание |
|------|----------|
| `Interface.cs` | Реализует `IExternalApplication`, создаёт вкладку и кнопку на ленте Revit. |
| `OpeningTask.addin` | XML-манифест для регистрации плагина в Revit. |

## Алгоритм работы

1. **Выбор моделей** — пользователь выбирает связанные модели MEP и АР/КР
2. **Фильтрация** — опционально фильтрует элементы по типам и параметрам
3. **Поиск пересечений** — `IntersectionService` находит все пересечения через Boolean операции
4. **Размещение боксов** — `CuboidPlacementService` создаёт семейства с правильной ориентацией
5. **Отчёт** — показывает количество созданных боксов и найденные дубликаты

## Лицензия

MIT
