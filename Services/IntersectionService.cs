using Autodesk.Revit.DB;
using OpeningTask.Helpers;
using OpeningTask.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpeningTask.Services
{
    /// <summary>
    /// Сервис для поиска пересечений элементов
    /// </summary>
    public class IntersectionService
    {
        private readonly Document _doc;

        public IntersectionService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Найти все пересечения MEP элементов со стенами и перекрытиями
        /// </summary>
        public List<IntersectionInfo> FindIntersections(
            IEnumerable<LinkedElementInfo> mepElements,
            IEnumerable<LinkedElementInfo> wallElements,
            IEnumerable<LinkedElementInfo> floorElements)
        {
            var intersections = new List<IntersectionInfo>();

            // Группируем элементы по связанным моделям для оптимизации
            var mepByLink = mepElements.GroupBy(e => e.LinkInstance.Id);
            var wallsByLink = wallElements.GroupBy(e => e.LinkInstance.Id);
            var floorsByLink = floorElements.GroupBy(e => e.LinkInstance.Id);

            // Обрабатываем каждый MEP элемент
            foreach (var mepInfo in mepElements)
            {
                var mepElement = mepInfo.LinkedElement;
                var mepTransform = mepInfo.LinkInstance.GetTotalTransform();

                // Получаем геометрию MEP элемента в координатах текущего документа
                var mepSolids = GetTransformedSolids(mepElement, mepTransform);
                if (!mepSolids.Any()) continue;

                // Проверяем пересечения со стенами
                foreach (var wallInfo in wallElements)
                {
                    var intersection = FindElementIntersection(
                        mepInfo, mepSolids, mepTransform,
                        wallInfo, HostElementType.Wall);

                    if (intersection != null)
                        intersections.Add(intersection);
                }

                // Проверяем пересечения с перекрытиями
                foreach (var floorInfo in floorElements)
                {
                    var intersection = FindElementIntersection(
                        mepInfo, mepSolids, mepTransform,
                        floorInfo, HostElementType.Floor);

                    if (intersection != null)
                        intersections.Add(intersection);
                }
            }

            return intersections;
        }

        /// <summary>
        /// Найти пересечение между MEP элементом и элементом вставки
        /// </summary>
        private IntersectionInfo FindElementIntersection(
            LinkedElementInfo mepInfo,
            List<Solid> mepSolids,
            Transform mepTransform,
            LinkedElementInfo hostInfo,
            HostElementType hostType)
        {
            var hostElement = hostInfo.LinkedElement;
            var hostTransform = hostInfo.LinkInstance.GetTotalTransform();

            // Получаем геометрию элемента вставки
            var hostSolids = GetTransformedSolids(hostElement, hostTransform);
            if (!hostSolids.Any()) return null;

            // Проверяем пересечение Solid-ов
            foreach (var mepSolid in mepSolids)
            {
                foreach (var hostSolid in hostSolids)
                {
                    try
                    {
                        // Используем BooleanOperationsUtils для проверки пересечения
                        var intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            mepSolid, hostSolid, BooleanOperationsType.Intersect);

                        if (intersectionSolid != null && intersectionSolid.Volume > 0.0001)
                        {
                            // Нашли пересечение - вычисляем параметры
                            return CreateIntersectionInfo(
                                mepInfo, mepTransform,
                                hostInfo, hostTransform, hostType,
                                intersectionSolid);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        // Пропускаем некорректные геометрии
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Создать информацию о пересечении
        /// </summary>
        private IntersectionInfo CreateIntersectionInfo(
            LinkedElementInfo mepInfo,
            Transform mepTransform,
            LinkedElementInfo hostInfo,
            Transform hostTransform,
            HostElementType hostType,
            Solid intersectionSolid)
        {
            var mepElement = mepInfo.LinkedElement;
            var hostElement = hostInfo.LinkedElement;

            // Вычисляем центр пересечения
            var centroid = intersectionSolid.ComputeCentroid();

            // Определяем тип и параметры MEP элемента
            var mepType = GetMepElementType(mepElement);
            var sectionType = GetMepSectionType(mepElement);
            var (width, height, diameter) = GetMepDimensions(mepElement);

            // Определяем направления
            var mepDirection = GetMepDirection(mepElement, mepTransform);
            var hostNormal = GetHostNormal(hostElement, hostTransform, hostType);

            // Толщина элемента вставки
            var hostThickness = GetHostThickness(hostElement, hostType);

            return new IntersectionInfo
            {
                MepElement = mepElement,
                MepLinkInstance = mepInfo.LinkInstance,
                HostElement = hostElement,
                HostLinkInstance = hostInfo.LinkInstance,
                InsertionPoint = centroid,
                MepDirection = mepDirection,
                HostNormal = hostNormal,
                MepType = mepType,
                SectionType = sectionType,
                HostType = hostType,
                MepWidth = width,
                MepHeight = height,
                MepDiameter = diameter,
                HostThickness = hostThickness
            };
        }

        /// <summary>
        /// Получить трансформированные Solid элемента
        /// </summary>
        private List<Solid> GetTransformedSolids(Element element, Transform transform)
        {
            var solids = new List<Solid>();

            var options = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            var geometry = element.get_Geometry(options);
            if (geometry == null) return solids;

            foreach (var geoObj in geometry)
            {
                solids.AddRange(ExtractSolids(geoObj, transform));
            }

            return solids.Where(s => s != null && s.Volume > 0).ToList();
        }

        /// <summary>
        /// Извлечь Solid из геометрического объекта
        /// </summary>
        private IEnumerable<Solid> ExtractSolids(GeometryObject geoObj, Transform transform)
        {
            if (geoObj is Solid solid && solid.Volume > 0)
            {
                yield return SolidUtils.CreateTransformed(solid, transform);
            }
            else if (geoObj is GeometryInstance geoInstance)
            {
                var instanceTransform = transform.Multiply(geoInstance.Transform);
                foreach (var instObj in geoInstance.GetInstanceGeometry())
                {
                    foreach (var s in ExtractSolids(instObj, instanceTransform))
                    {
                        yield return s;
                    }
                }
            }
        }

        /// <summary>
        /// Определить тип MEP элемента
        /// </summary>
        private MepElementType GetMepElementType(Element element)
        {
            if (element.Category == null) return MepElementType.Unknown;

            var categoryId = element.Category.Id.Value;

            if (categoryId == (long)BuiltInCategory.OST_PipeCurves ||
                categoryId == (long)BuiltInCategory.OST_FlexPipeCurves)
                return MepElementType.Pipe;

            if (categoryId == (long)BuiltInCategory.OST_DuctCurves ||
                categoryId == (long)BuiltInCategory.OST_FlexDuctCurves)
                return MepElementType.Duct;

            if (categoryId == (long)BuiltInCategory.OST_CableTray ||
                categoryId == (long)BuiltInCategory.OST_Conduit)
                return MepElementType.Tray;

            return MepElementType.Unknown;
        }

        /// <summary>
        /// Определить тип сечения MEP элемента
        /// </summary>
        private MepSectionType GetMepSectionType(Element element)
        {
            // Проверяем параметр диаметра
            var diameterParam = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) ??
                               element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);

            if (diameterParam != null && diameterParam.HasValue)
                return MepSectionType.Round;

            // Проверяем параметры ширины/высоты воздуховода
            var widthParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var heightParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);

            if (widthParam != null && heightParam != null)
                return MepSectionType.Rectangular;

            // По умолчанию - прямоугольное (для лотков)
            return MepSectionType.Rectangular;
        }

        /// <summary>
        /// Получить размеры MEP элемента
        /// </summary>
        private (double width, double height, double diameter) GetMepDimensions(Element element)
        {
            double width = 0, height = 0, diameter = 0;

            // Диаметр для труб
            var diameterParam = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) ??
                               element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (diameterParam != null && diameterParam.HasValue)
            {
                diameter = diameterParam.AsDouble();
                width = diameter;
                height = diameter;
            }

            // Ширина и высота для прямоугольных элементов
            var widthParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var heightParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);

            if (widthParam != null && widthParam.HasValue)
                width = widthParam.AsDouble();
            if (heightParam != null && heightParam.HasValue)
                height = heightParam.AsDouble();

            // Для кабельных лотков
            var trayWidthParam = element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            var trayHeightParam = element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);

            if (trayWidthParam != null && trayWidthParam.HasValue)
                width = trayWidthParam.AsDouble();
            if (trayHeightParam != null && trayHeightParam.HasValue)
                height = trayHeightParam.AsDouble();

            return (width, height, diameter);
        }

        /// <summary>
        /// Получить направление MEP элемента
        /// </summary>
        private XYZ GetMepDirection(Element element, Transform transform)
        {
            if (element.Location is LocationCurve locationCurve)
            {
                var curve = locationCurve.Curve;
                var direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                return transform.OfVector(direction);
            }

            return XYZ.BasisX;
        }

        /// <summary>
        /// Получить нормаль элемента вставки
        /// </summary>
        private XYZ GetHostNormal(Element element, Transform transform, HostElementType hostType)
        {
            if (hostType == HostElementType.Wall && element is Wall wall)
            {
                var locationCurve = wall.Location as LocationCurve;
                if (locationCurve != null)
                {
                    var curve = locationCurve.Curve;
                    var direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                    // Нормаль стены перпендикулярна направлению
                    var normal = new XYZ(-direction.Y, direction.X, 0);
                    return transform.OfVector(normal);
                }
            }

            if (hostType == HostElementType.Floor)
            {
                // Нормаль перекрытия - вертикаль
                return transform.OfVector(XYZ.BasisZ);
            }

            return XYZ.BasisZ;
        }

        /// <summary>
        /// Получить толщину элемента вставки
        /// </summary>
        private double GetHostThickness(Element element, HostElementType hostType)
        {
            if (hostType == HostElementType.Wall && element is Wall wall)
            {
                return wall.Width;
            }

            if (hostType == HostElementType.Floor && element is Floor floor)
            {
                var floorType = floor.Document.GetElement(floor.GetTypeId()) as FloorType;
                if (floorType != null)
                {
                    var compound = floorType.GetCompoundStructure();
                    if (compound != null)
                    {
                        return compound.GetWidth();
                    }
                }
            }

            return 0;
        }
    }
}