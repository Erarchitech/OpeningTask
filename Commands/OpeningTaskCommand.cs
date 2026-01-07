using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OpeningTask.ViewModels;
using System;

namespace OpeningTask.Commands
{
    /// <summary>
    /// Основная команда плагина "Задание на отверстия"
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class OpeningTaskCommand : IExternalCommand
    {
        /// <summary>
        /// Путь к семействам кубиков
        /// </summary>
        public const string CuboidFamiliesPath = @"C:\OpeningTaskResources";

        /// <summary>
        /// Имена семейств кубиков
        /// </summary>
        public static class CuboidFamilyNames
        {
            public const string FloorRound = "Кубик_Перекрытие_Круг";
            public const string FloorRectangular = "Кубик_Перекрытие_Прямоугольный";
            public const string WallRound = "Кубик_Стена_Круг";
            public const string WallRectangular = "Кубик_Стена_Прямоугольный";
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Проверка наличия активного документа
                if (doc == null)
                {
                    TaskDialog.Show("Ошибка", "Нет активного документа.");
                    return Result.Failed;
                }

                // Проверка, что документ не является семейством
                if (doc.IsFamilyDocument)
                {
                    TaskDialog.Show("Ошибка", "Команда не может быть выполнена в редакторе семейств.");
                    return Result.Failed;
                }

                // Создание ViewModel и отображение главного окна
                var viewModel = new MainWindowViewModel(uiDoc);
                var mainWindow = new Views.MainWindow(viewModel);

                // Показываем окно как диалог
                bool? dialogResult = mainWindow.ShowDialog();

                if (dialogResult == true)
                {
                    // Пользователь нажал OK - здесь будет логика второго этапа
                    // (создание кубиков в местах пересечений)
                    TaskDialog.Show("Информация", 
                        "Настройки фильтрации применены.\n\n" +
                        "Функционал создания кубиков будет реализован на втором этапе.");
                    return Result.Succeeded;
                }
                else
                {
                    // Пользователь отменил операцию
                    return Result.Cancelled;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = $"Произошла ошибка: {ex.Message}\n\n{ex.StackTrace}";
                TaskDialog.Show("Ошибка", message);
                return Result.Failed;
            }
        }
    }
}
