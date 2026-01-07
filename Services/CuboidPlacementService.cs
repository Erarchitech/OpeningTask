using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using OpeningTask.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpeningTask.Services
{
    /// <summary>
    /// Сервис для размещения кубиков в местах пересечений
    /// </summary>
    public class CuboidPlacementService
    {
        private readonly Document _doc;
        private readonly CuboidSettings _settings;
        private readonly Dictionary<string, FamilySymbol> _loadedSymbols;

        // Конвертация мм в футы
        private const double MmToFeet = 1.0 / 304.8;

        public CuboidPlacementService(Document doc, CuboidSettings settings)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _loadedSymbols = new Dictionary<string, FamilySymbol>();
        }

        /// <summary>
        /// Разместить кубики в местах пересечений
        /// </summary>
        public List<FamilyInstance> PlaceCuboids(IEnumerable<IntersectionInfo> intersections)
        {
            var placedInstances = new List<FamilyInstance>();

            using (var transaction = new Transaction(_doc, "Размещение кубиков"))
            {
                transaction.Start();

                foreach (var intersection in intersections)
                {
                    try
                    {
                        var instance = PlaceSingleCuboid(intersection);
                        if (instance != null)
                        {
                            placedInstances.Add(instance);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Логируем ошибку и продолжаем
                        System.Diagnostics.Debug.WriteLine($"Error placing cuboid: {ex.Message}");
                    }
                }

                transaction.Commit();
            }

            return placedInstances;
        }

        /// <summary>
        /// Разместить один кубик
        /// </summary>
        private FamilyInstance PlaceSingleCuboid(IntersectionInfo intersection)
        {
            // Получаем подходящий типоразмер семейства
            var symbol = GetOrLoadFamilySymbol(intersection);
            if (symbol == null) return null;

            // Активируем типоразмер если нужно
            if (!symbol.IsActive)
                symbol.Activate();

            // Вычисляем параметры кубика
            var cuboidParams = CalculateCuboidParameters(intersection);

            // Размещаем экземпляр семейства
            FamilyInstance instance;

            if (intersection.HostType == HostElementType.Wall)
            {
                // Для стен используем размещение по точке с ориентацией
                instance = _doc.Create.NewFamilyInstance(
                    intersection.InsertionPoint,
                    symbol,
                    StructuralType.NonStructural);
            }
            else
            {
                // Для перекрытий - аналогично
                instance = _doc.Create.NewFamilyInstance(
                    intersection.InsertionPoint,
                    symbol,
                    StructuralType.NonStructural);
            }

            if (instance == null) return null;

            // Ориентируем кубик
            OrientCuboid(instance, intersection);

            // Заполняем параметры
            SetCuboidParameters(instance, cuboidParams, intersection);

            return instance;
        }

        /// <summary>
        /// Получить или загрузить типоразмер семейства
        /// </summary>
        private FamilySymbol GetOrLoadFamilySymbol(IntersectionInfo intersection)
        {
            var familyName = _settings.GetCuboidFamilyName(
                intersection.HostType,
                intersection.SectionType,
                intersection.MepType);

            // Проверяем кэш
            if (_loadedSymbols.TryGetValue(familyName, out var cachedSymbol))
                return cachedSymbol;

            // Ищем уже загруженное семейство
            var existingSymbol = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(s => s.Family.Name == familyName);

            if (existingSymbol != null)
            {
                _loadedSymbols[familyName] = existingSymbol;
                return existingSymbol;
            }

            // Загружаем семейство
            var familyPath = _settings.GetCuboidFamilyPath(
                intersection.HostType,
                intersection.SectionType,
                intersection.MepType);

            if (!File.Exists(familyPath))
            {
                throw new FileNotFoundException($"Family file not found: {familyPath}");
            }

            if (_doc.LoadFamily(familyPath, out Family family))
            {
                var symbolId = family.GetFamilySymbolIds().FirstOrDefault();
                if (symbolId != null && symbolId != ElementId.InvalidElementId)
                {
                    var symbol = _doc.GetElement(symbolId) as FamilySymbol;
                    _loadedSymbols[familyName] = symbol;
                    return symbol;
                }
            }

            return null;
        }

        /// <summary>
        /// Вычислить параметры кубика
        /// </summary>
        private CuboidParameters CalculateCuboidParameters(IntersectionInfo intersection)
        {
            var result = new CuboidParameters();

            // Конвертируем настройки из мм в футы
            double minOffsetFeet = _settings.MinOffset * MmToFeet;
            double protrusionFeet = _settings.Protrusion * MmToFeet;
            double roundingFeet = _settings.RoundingValue * MmToFeet;

            // Вычисляем размеры на основе сечения MEP элемента
            if (intersection.SectionType == MepSectionType.Round)
            {
                // Для круглого сечения
                double size = intersection.MepDiameter + 2 * minOffsetFeet;
                size = RoundUpToNearest(size, roundingFeet);

                result.Width = size;
                result.Height = size;
                result.Diameter = size;
            }
            else
            {
                // Для прямоугольного сечения
                double width = intersection.MepWidth + 2 * minOffsetFeet;
                double height = intersection.MepHeight + 2 * minOffsetFeet;

                result.Width = RoundUpToNearest(width, roundingFeet);
                result.Height = RoundUpToNearest(height, roundingFeet);
            }

            // Толщина кубика = толщина элемента вставки + выступ с обеих сторон
            result.Thickness = intersection.HostThickness + 2 * protrusionFeet;
            result.Protrusion = protrusionFeet;

            return result;
        }

        /// <summary>
        /// Округление вверх до ближайшего значения
        /// </summary>
        private double RoundUpToNearest(double value, double rounding)
        {
            if (rounding <= 0) return value;
            return Math.Ceiling(value / rounding) * rounding;
        }

        /// <summary>
        /// Ориентировать кубик
        /// </summary>
        private void OrientCuboid(FamilyInstance instance, IntersectionInfo intersection)
        {
            // Получаем текущую ориентацию
            var currentFacing = instance.FacingOrientation;
            var targetNormal = intersection.HostNormal;

            if (intersection.HostType == HostElementType.Wall)
            {
                // Для стен - ориентируем по нормали стены
                var angle = GetAngleBetweenVectors(currentFacing, targetNormal);
                if (Math.Abs(angle) > 0.001)
                {
                    var axis = Line.CreateBound(
                        intersection.InsertionPoint,
                        intersection.InsertionPoint + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                }
            }
            else
            {
                // Для перекрытий - ориентируем по направлению MEP элемента
                var mepDirHorizontal = new XYZ(
                    intersection.MepDirection.X,
                    intersection.MepDirection.Y,
                    0).Normalize();

                if (mepDirHorizontal.GetLength() > 0.001)
                {
                    var angle = GetAngleBetweenVectors(XYZ.BasisX, mepDirHorizontal);
                    var axis = Line.CreateBound(
                        intersection.InsertionPoint,
                        intersection.InsertionPoint + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                }
            }
        }

        /// <summary>
        /// Получить угол между векторами
        /// </summary>
        private double GetAngleBetweenVectors(XYZ v1, XYZ v2)
        {
            var cross = v1.CrossProduct(v2);
            var dot = v1.DotProduct(v2);
            return Math.Atan2(cross.Z, dot);
        }

        /// <summary>
        /// Установить параметры кубика
        /// </summary>
        private void SetCuboidParameters(FamilyInstance instance, CuboidParameters cuboidParams, IntersectionInfo intersection)
        {
            // Ширина
            SetParameterByGuid(instance, CuboidSettings.WidthParamGuid, cuboidParams.Width);

            // Высота
            SetParameterByGuid(instance, CuboidSettings.HeightParamGuid, cuboidParams.Height);

            // Толщина (зависит от типа элемента вставки)
            if (intersection.HostType == HostElementType.Wall)
            {
                SetParameterByGuid(instance, CuboidSettings.WallThicknessParamGuid, intersection.HostThickness);
            }
            else
            {
                SetParameterByGuid(instance, CuboidSettings.FloorThicknessParamGuid, intersection.HostThickness);
            }

            // Дополнительная толщина (выступ)
            SetParameterByGuid(instance, CuboidSettings.AdditionalThickness1ParamGuid, cuboidParams.Protrusion);
            SetParameterByGuid(instance, CuboidSettings.AdditionalThickness2ParamGuid, cuboidParams.Protrusion);
        }

        /// <summary>
        /// Установить параметр по GUID
        /// </summary>
        private void SetParameterByGuid(Element element, Guid paramGuid, double value)
        {
            var param = element.get_Parameter(paramGuid);
            if (param != null && !param.IsReadOnly)
            {
                param.Set(value);
            }
        }

        /// <summary>
        /// Параметры кубика
        /// </summary>
        private class CuboidParameters
        {
            public double Width { get; set; }
            public double Height { get; set; }
            public double Diameter { get; set; }
            public double Thickness { get; set; }
            public double Protrusion { get; set; }
        }
    }
}