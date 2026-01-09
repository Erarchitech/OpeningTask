using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OpeningTask.Helpers;
using OpeningTask.Models;
using OpeningTask.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace OpeningTask.Helpers
{
    /// <summary>
    /// Данные запроса на размещение кубиков
    /// </summary>
    public class CuboidPlacementRequest
    {
        public List<LinkedElementInfo> MepElements { get; set; }
        public List<LinkedElementInfo> WallElements { get; set; }
        public List<LinkedElementInfo> FloorElements { get; set; }
        public CuboidSettings Settings { get; set; }
    }

    /// <summary>
    /// Обработчик внешнего события для размещения кубиков
    /// </summary>
    public class CuboidPlacementEventHandler : IExternalEventHandler
    {
        /// <summary>
        /// Текущий запрос на размещение
        /// </summary>
        public CuboidPlacementRequest Request { get; set; }

        /// <summary>
        /// Результат выполнения (количество созданных кубиков)
        /// </summary>
        public int ResultCount { get; private set; }

        /// <summary>
        /// Сообщение об ошибке (если есть)
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Успешно ли выполнено
        /// </summary>
        public bool IsSuccess { get; private set; }

        /// <summary>
        /// ID существующих кубиков, из-за которых создание было пропущено (дубликаты)
        /// </summary>
        public List<ElementId> DuplicateCuboidIds { get; private set; } = new List<ElementId>();

        /// <summary>
        /// Событие завершения операции
        /// </summary>
        public event Action<bool, int, string> OperationCompleted;

        public Dispatcher UiDispatcher { get; set; }

        public void Execute(UIApplication app)
        {
            ResultCount = 0;
            ErrorMessage = null;
            IsSuccess = false;
            DuplicateCuboidIds.Clear();

            RevitTrace.Info("ExternalEvent.Execute: start");

            if (Request == null)
            {
                ErrorMessage = "Запрос не задан";
                RevitTrace.Warn("ExternalEvent.Execute: Request is null");
                RaiseOperationCompleted(false, 0, ErrorMessage);
                return;
            }

            try
            {
                var doc = app.ActiveUIDocument.Document;

                RevitTrace.Info($"ExternalEvent.Execute: request mep={Request.MepElements?.Count ?? 0}, walls={Request.WallElements?.Count ?? 0}, floors={Request.FloorElements?.Count ?? 0}");

                // Обновляем ссылки на элементы перед использованием
                var refreshedMepElements = RefreshElementReferences(doc, Request.MepElements);
                var refreshedWallElements = RefreshElementReferences(doc, Request.WallElements);
                var refreshedFloorElements = RefreshElementReferences(doc, Request.FloorElements);

                if (!refreshedMepElements.Any())
                {
                    IsSuccess = true;
                    ErrorMessage = "Не найдено действительных MEP элементов";
                    RevitTrace.Warn("ExternalEvent.Execute: no valid MEP elements after refresh");
                    RaiseOperationCompleted(true, 0, ErrorMessage);
                    return;
                }

                // Поиск пересечений
                var intersectionService = new IntersectionService(doc);
                var intersections = intersectionService.FindIntersections(
                    refreshedMepElements,
                    refreshedWallElements,
                    refreshedFloorElements);

                if (intersections.Count == 0)
                {
                    IsSuccess = true;
                    ErrorMessage = "Пересечения не найдены";
                    RevitTrace.Info("ExternalEvent.Execute: intersections=0");
                    RaiseOperationCompleted(true, 0, ErrorMessage);
                    return;
                }

                RevitTrace.Info($"ExternalEvent.Execute: intersections={intersections.Count}");

                // Размещение кубиков
                var placementService = new CuboidPlacementService(doc, Request.Settings);
                var placedCuboids = placementService.PlaceCuboids(intersections);

                try
                {
                    var dupIds = new HashSet<long>();
                    foreach (var kvp in placementService.DuplicateExistingIds)
                    {
                        foreach (var id in kvp.Value)
                        {
                            if (id != null)
                                dupIds.Add(id.Value);
                        }
                    }

                    DuplicateCuboidIds = dupIds.Select(v => new ElementId(v)).ToList();
                    RevitTrace.Info($"ExternalEvent.Execute: duplicatesExistingIds={DuplicateCuboidIds.Count}");
                }
                catch (Exception ex)
                {
                    DuplicateCuboidIds = new List<ElementId>();
                    RevitTrace.Warn($"ExternalEvent.Execute: failed to collect duplicate ids: {ex.Message}");
                }

                ResultCount = placedCuboids.Count;
                IsSuccess = true;
                RevitTrace.Info($"ExternalEvent.Execute: placed={ResultCount}");
                RaiseOperationCompleted(true, ResultCount, null);
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                IsSuccess = false;
                RevitTrace.Error("ExternalEvent.Execute: exception", ex);
                RaiseOperationCompleted(false, 0, ErrorMessage);
            }
        }

        private void RaiseOperationCompleted(bool success, int count, string errorMessage)
        {
            var handler = OperationCompleted;
            if (handler == null)
                return;

            var dispatcher = UiDispatcher ?? Application.Current?.Dispatcher;
            if (dispatcher != null)
            {
                if (dispatcher.CheckAccess())
                    handler(success, count, errorMessage);
                else
                    dispatcher.BeginInvoke(new Action(() => handler(success, count, errorMessage)));
                return;
            }

            handler(success, count, errorMessage);
        }

        /// <summary>
        /// Обновить ссылки на элементы из связанных моделей
        /// </summary>
        private List<LinkedElementInfo> RefreshElementReferences(Document doc, List<LinkedElementInfo> elements)
        {
            var result = new List<LinkedElementInfo>();

            if (doc == null)
                return result;

            if (elements == null)
                return result;

            foreach (var info in elements)
            {
                try
                {
                    if (info == null)
                        continue;

                    if (info.LinkInstance == null || !info.LinkInstance.IsValidObject)
                        continue;

                    if (info.LinkedElementId == null || info.LinkedElementId == ElementId.InvalidElementId)
                        continue;

                    // Получаем актуальный LinkInstance по ID из текущего документа
                    var linkInstance = doc.GetElement(info.LinkInstance.Id) as RevitLinkInstance;
                    if (linkInstance == null || !linkInstance.IsValidObject)
                        continue;

                    // Получаем актуальный документ связанной модели
                    var linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null)
                        continue;

                    // Получаем актуальный элемент по ID
                    var element = linkedDoc.GetElement(info.LinkedElementId);
                    if (element == null || !element.IsValidObject)
                        continue;

                    // Создаём обновлённую информацию
                    result.Add(new LinkedElementInfo
                    {
                        LinkInstance = linkInstance,
                        LinkedElementId = info.LinkedElementId,
                        LinkedElement = element,
                        LinkedDocument = linkedDoc
                    });
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    // Пропускаем недействительные элементы
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Пропускаем недействительные элементы
                }
            }

            return result;
        }

        public string GetName()
        {
            return "Cuboid Placement Handler";
        }
    }
}
