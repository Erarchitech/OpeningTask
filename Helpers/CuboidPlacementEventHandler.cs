using Autodesk.Revit.UI;
using OpeningTask.Helpers;
using OpeningTask.Models;
using OpeningTask.Services;
using System;
using System.Collections.Generic;
using System.Windows;

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
        /// Событие завершения операции
        /// </summary>
        public event Action<bool, int, string> OperationCompleted;

        public void Execute(UIApplication app)
        {
            ResultCount = 0;
            ErrorMessage = null;
            IsSuccess = false;

            if (Request == null)
            {
                ErrorMessage = "Запрос не задан";
                OperationCompleted?.Invoke(false, 0, ErrorMessage);
                return;
            }

            try
            {
                var doc = app.ActiveUIDocument.Document;

                // Поиск пересечений
                var intersectionService = new IntersectionService(doc);
                var intersections = intersectionService.FindIntersections(
                    Request.MepElements,
                    Request.WallElements,
                    Request.FloorElements);

                if (intersections.Count == 0)
                {
                    IsSuccess = true;
                    ErrorMessage = "Пересечения не найдены";
                    OperationCompleted?.Invoke(true, 0, ErrorMessage);
                    return;
                }

                // Размещение кубиков
                var placementService = new CuboidPlacementService(doc, Request.Settings);
                var placedCuboids = placementService.PlaceCuboids(intersections);

                ResultCount = placedCuboids.Count;
                IsSuccess = true;
                OperationCompleted?.Invoke(true, ResultCount, null);
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                IsSuccess = false;
                OperationCompleted?.Invoke(false, 0, ErrorMessage);
            }
        }

        public string GetName()
        {
            return "Cuboid Placement Handler";
        }
    }
}
