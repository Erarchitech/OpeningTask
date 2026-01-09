using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using OpeningTask.Helpers;
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

            // Корректируем точку вставки с учетом того, что точка вставки семейства на верхней грани
            var adjustedInsertionPoint = CalculateAdjustedInsertionPoint(
                intersection, cuboidParams);

            RevitTrace.Info(
                $"PlaceSingleCuboid: intersectionPoint=({intersection.InsertionPoint.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{intersection.InsertionPoint.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{intersection.InsertionPoint.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                $"adjustedInsertionPoint=({adjustedInsertionPoint.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{adjustedInsertionPoint.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{adjustedInsertionPoint.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                $"hostType={intersection.HostType} sectionType={intersection.SectionType}");

            // Размещаем экземпляр семейства
            FamilyInstance instance = _doc.Create.NewFamilyInstance(
                adjustedInsertionPoint,
                symbol,
                StructuralType.NonStructural);

            if (instance == null) return null;

            try
            {
                _doc.Regenerate();
                var lp = instance.Location as LocationPoint;
                if (lp != null)
                {
                    var p = lp.Point;
                    RevitTrace.Info(
                        $"PlaceSingleCuboid: instanceLocation=({p.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{p.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{p.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                        $"delta=({(p.X - adjustedInsertionPoint.X).ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{(p.Y - adjustedInsertionPoint.Y).ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{(p.Z - adjustedInsertionPoint.Z).ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");
                }
                else
                {
                    RevitTrace.Warn("PlaceSingleCuboid: instance.Location is not LocationPoint after Regenerate");
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("PlaceSingleCuboid: failed to read instance location after creation", ex);
            }

            // ВАЖНО: сначала заполняем параметры, чтобы геометрия была правильной до поворота
            SetCuboidParameters(instance, cuboidParams, intersection);

            // Ориентируем кубик (передаём точку размещения для оси вращения)
            OrientCuboid(instance, intersection, cuboidParams, adjustedInsertionPoint);

            return instance;
        }

        /// <summary>
        /// Вычислить скорректированную точку вставки.
        /// Точка вставки семейства находится в центре кубика в плане.
        /// Для перекрытий смещаем по Z на половину толщины хоста.
        /// </summary>
        private XYZ CalculateAdjustedInsertionPoint(IntersectionInfo intersection, CuboidParameters cuboidParams)
        {
            var centerPoint = intersection.InsertionPoint;

            if (intersection.HostType == HostElementType.Floor)
            {
                // Для перекрытий: смещаем по Z на половину толщины перекрытия
                double halfHostThickness = intersection.HostThickness / 2.0;
                return new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z + halfHostThickness);
            }

            if (intersection.HostType == HostElementType.Wall)
            {
                // Для стен точка вставки семейства прямоугольного кубика находится
                // в центре нижней грани кубика. Поэтому опускаем точку вставки вниз
                // на половину высоты САМОГО кубика (включая отступы/округления),
                // т.к. cuboidParams уже рассчитаны с учетом MinOffset и rounding.
                double halfCuboidHeight;
                if (intersection.SectionType == MepSectionType.Round)
                {
                    halfCuboidHeight = (cuboidParams?.Diameter ?? 0) / 2.0;
                }
                else
                {
                    halfCuboidHeight = (cuboidParams?.Height ?? 0) / 2.0;
                }

                if (halfCuboidHeight > 1e-9)
                {
                    return new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z - halfCuboidHeight);
                }

                return centerPoint;
            }

            // Fallback
            return centerPoint;
        }

        private XYZ ApplyFloorRectangularInsertionCompensation(IntersectionInfo intersection, XYZ insertionPoint, CuboidParameters cuboidParams)
        {
            // Наблюдаемое смещение равно половине разницы сторон (в плане).
            // При текущем семейства: локальная X (HandOrientation) указывает вдоль "Width".
            var delta = (cuboidParams.Width - cuboidParams.Height) / 2.0;
            if (Math.Abs(delta) < 1e-9)
                return insertionPoint;

            // Компенсацию делаем вдоль оси сечения MEP (BasisX коннектора), спроецированной в XY.
            // Это обеспечивает корректность независимо от ориентации в плане.
            var sectionXAxis = GetMepSectionXAxis(intersection.MepElement);
            if (sectionXAxis == null)
            {
                RevitTrace.Warn("ApplyFloorRectangularInsertionCompensation: sectionXAxis is null, skipping");
                return insertionPoint;
            }

            var dir = new XYZ(sectionXAxis.X, sectionXAxis.Y, 0);
            if (dir.GetLength() < 1e-6)
            {
                RevitTrace.Warn("ApplyFloorRectangularInsertionCompensation: sectionXAxis is vertical/zero in XY, skipping");
                return insertionPoint;
            }
            dir = dir.Normalize();

            // Сдвигаем точку вставки вдоль dir на delta.
            // Знак выбран так, чтобы после поворота кубик совпал с точкой пересечения.
            var compensated = insertionPoint + dir.Multiply(delta);

            RevitTrace.Info(
                $"ApplyFloorRectangularInsertionCompensation: delta={(delta * 304.8).ToString("F3", System.Globalization.CultureInfo.InvariantCulture)}mm " +
                $"dir=({dir.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{dir.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                $"from=({insertionPoint.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{insertionPoint.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{insertionPoint.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                $"to=({compensated.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{compensated.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{compensated.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");

            return compensated;
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

            RevitTrace.Info($"GetOrLoadFamilySymbol: familyName={familyName}, hostType={intersection.HostType}, section={intersection.SectionType}, mepType={intersection.MepType}");

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
                throw new FileNotFoundException($"Файл семейства не найден: {familyPath}");
            }

            if (_doc.LoadFamily(familyPath, out Family family))
            {
                RevitTrace.Info($"GetOrLoadFamilySymbol: family loaded: {family.Name}");
                var symbolId = family.GetFamilySymbolIds().FirstOrDefault();
                if (symbolId != null && symbolId != ElementId.InvalidElementId)
                {
                    var symbol = _doc.GetElement(symbolId) as FamilySymbol;
                    _loadedSymbols[familyName] = symbol;
                    return symbol;
                }
            }
            else
            {
                RevitTrace.Warn($"GetOrLoadFamilySymbol: LoadFamily returned false: {familyPath}");
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
        private void OrientCuboid(FamilyInstance instance, IntersectionInfo intersection, CuboidParameters cuboidParams, XYZ insertionPoint)
        {
            // Используем переданную точку размещения для оси вращения
            // НЕ берём instance.Location — он может быть (0,0,0) до commit транзакции
            var instanceLocation = insertionPoint;
            
            // Диагностика координат
            RevitTrace.Info($"OrientCuboid coords: intersectionPoint=({intersection.InsertionPoint.X:F4},{intersection.InsertionPoint.Y:F4},{intersection.InsertionPoint.Z:F4}) " +
                $"insertionPoint=({insertionPoint.X:F4},{insertionPoint.Y:F4},{insertionPoint.Z:F4}) " +
                $"rotationAxis=({instanceLocation.X:F4},{instanceLocation.Y:F4},{instanceLocation.Z:F4}) " +
                $"hostType={intersection.HostType}");
            
            // Получаем текущую ориентацию
            var currentFacing = instance.FacingOrientation;
            var targetNormal = intersection.HostNormal;

            if (intersection.HostType == HostElementType.Wall)
            {
                // Для стен - базовая ориентация по нормали стены
                var angle = GetAngleBetweenVectors(currentFacing, targetNormal);
                if (Math.Abs(angle) > 0.001)
                {
                    var axis = Line.CreateBound(
                        instanceLocation,
                        instanceLocation + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                }

                // Для прямоугольных сечений: корректируем ориентацию по реальному сечению MEP
                if (intersection.SectionType == MepSectionType.Rectangular && cuboidParams != null)
                {
                    OrientCuboidForWallByMepSection(instance, intersection, cuboidParams, instanceLocation);
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
                        instanceLocation,
                        instanceLocation + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                    
                    // После поворота возвращаем экземпляр в правильную позицию
                    ResetInstanceLocation(instance, instanceLocation);
                }
            }

            // Коррекция для прямоугольных кубиков в ПЕРЕКРЫТИЯХ: сторона "Ширина" должна быть вдоль направления MEP.
            // Для СТЕН эта коррекция не нужна — кубик уже правильно ориентирован по нормали стены.
            // Иначе часть элементов поворачивается на 90° из-за различий осей семейства.
            if (intersection.HostType == HostElementType.Wall)
            {
                // Для стен коррекция уже сделана в OrientCuboidForWallByMepSection
                return;
            }

            try
            {
                RevitTrace.Info($"OrientCuboid rect-check: sectionType={intersection.SectionType} cuboidParams={(cuboidParams != null ? "ok" : "null")} mepDir=({intersection.MepDirection.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{intersection.MepDirection.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{intersection.MepDirection.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");

                if (intersection.SectionType != MepSectionType.Rectangular)
                    return;

                if (cuboidParams == null)
                    return;

                RevitTrace.Info($"OrientCuboid rect-params: width={cuboidParams.Width.ToString("G", System.Globalization.CultureInfo.InvariantCulture)} height={cuboidParams.Height.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}");

                var mepDir = new XYZ(intersection.MepDirection.X, intersection.MepDirection.Y, 0);
                if (mepDir.GetLength() < 1e-6)
                {
                    // MEP идёт вертикально через перекрытие — направление MEP не даёт информации о повороте.
                    // Извлекаем ориентацию сечения через коннекторы MEP-элемента.
                    RevitTrace.Info("OrientCuboid: mepDir is vertical, extracting section orientation from connectors");
                    
                    var sectionXAxis = GetMepSectionXAxis(intersection.MepElement);
                    if (sectionXAxis != null)
                    {
                        RevitTrace.Info($"OrientCuboid: sectionXAxis=({sectionXAxis.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{sectionXAxis.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{sectionXAxis.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");
                        
                        // Проецируем ось X сечения на горизонтальную плоскость
                        var sectionXHoriz = new XYZ(sectionXAxis.X, sectionXAxis.Y, 0);
                        if (sectionXHoriz.GetLength() > 1e-6)
                        {
                            sectionXHoriz = sectionXHoriz.Normalize();
                            
                            // Получаем текущую ось X семейства кубика
                            var handCurrent = XYZ.BasisX;
                            try { handCurrent = instance.HandOrientation; } catch { }
                            handCurrent = new XYZ(handCurrent.X, handCurrent.Y, 0);
                            if (handCurrent.GetLength() > 1e-6)
                            {
                                handCurrent = handCurrent.Normalize();
                                
                                // Вычисляем угол поворота от текущей оси к оси сечения
                                var crossZSec = handCurrent.X * sectionXHoriz.Y - handCurrent.Y * sectionXHoriz.X;
                                var dotVal = handCurrent.DotProduct(sectionXHoriz);
                                var angleVal = Math.Atan2(crossZSec, dotVal);
                                
                                if (Math.Abs(angleVal) > 1e-6)
                                {
                                    var vertAxis = Line.CreateBound(instanceLocation, instanceLocation + XYZ.BasisZ);
                                    ElementTransformUtils.RotateElement(_doc, instance.Id, vertAxis, angleVal);
                                    RevitTrace.Info($"OrientCuboid: rotated by section orientation angle={angleVal.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}");
                                    
                                    // После поворота возвращаем экземпляр в правильную позицию
                                    ResetInstanceLocation(instance, instanceLocation);
                                }
                            }
                        }
                    }
                    else
                    {
                        // Fallback: если не удалось получить ориентацию сечения
                        RevitTrace.Warn("OrientCuboid: could not extract section orientation, using fallback");
                        if (cuboidParams.Height > cuboidParams.Width + 1e-6)
                        {
                            var vertAxis = Line.CreateBound(instanceLocation, instanceLocation + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, instance.Id, vertAxis, Math.PI / 2.0);
                            RevitTrace.Info("OrientCuboid: rotated 90deg for vertical MEP (height > width)");
                            
                            // После поворота возвращаем экземпляр в правильную позицию
                            ResetInstanceLocation(instance, instanceLocation);
                        }
                    }
                    return;
                }
                mepDir = mepDir.Normalize();

                // Ось вращения - вертикаль через точку размещения экземпляра
                var axis = Line.CreateBound(instanceLocation, instanceLocation + XYZ.BasisZ);

                // Локальная ось X семейства в мире
                var hand = XYZ.BasisX;
                try { hand = instance.HandOrientation; } catch { hand = instance.FacingOrientation; }
                hand = new XYZ(hand.X, hand.Y, 0);
                if (hand.GetLength() < 1e-6)
                    return;
                hand = hand.Normalize();

                // Локальная ось Y семейства в мире (в плоскости)
                var handY = XYZ.BasisZ.CrossProduct(hand);
                handY = new XYZ(handY.X, handY.Y, 0);
                if (handY.GetLength() < 1e-6)
                    return;
                handY = handY.Normalize();

                // Какая ось семьи должна быть вдоль MEP
                // Предположение принятое в проекте: Width параметр вдоль локальной X (hand), Height/Length вдоль локальной Y.
                var wantWidthAlongMep = cuboidParams.Width >= cuboidParams.Height;
                var fromAxis = wantWidthAlongMep ? hand : handY;

                RevitTrace.Info(
                    $"OrientCuboid dbg: instanceId={instance.Id.Value} width={cuboidParams.Width.ToString("G", System.Globalization.CultureInfo.InvariantCulture)} height={cuboidParams.Height.ToString("G", System.Globalization.CultureInfo.InvariantCulture)} " +
                    $"mepDir=({mepDir.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{mepDir.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                    $"hand=({hand.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{hand.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                    $"handY=({handY.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{handY.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)}) " +
                    $"from={(wantWidthAlongMep ? "X" : "Y")}");

                // Точный угол поворота fromAxis -> mepDir в плоскости XY
                var crossZ = fromAxis.X * mepDir.Y - fromAxis.Y * mepDir.X;
                var dot = fromAxis.DotProduct(mepDir);
                var angle = Math.Atan2(crossZ, dot);

                if (Math.Abs(angle) < 1e-6)
                    return;

                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);

                RevitTrace.Info($"OrientCuboid: aligned rectangular instanceId={instance.Id.Value} angle={angle.ToString("G", System.Globalization.CultureInfo.InvariantCulture)} wantWidthAlongMep={wantWidthAlongMep}");

                // После поворота возвращаем экземпляр в правильную позицию
                ResetInstanceLocation(instance, instanceLocation);

                // Автокоррекция: если семейство считает "ширину" вдоль другой оси, то после выравнивания
                // ожидаемая ось (X для Width, Y для Height) может оказаться поперёк MEP.
                try
                {
                    var hand2 = new XYZ(instance.HandOrientation.X, instance.HandOrientation.Y, 0);
                    if (hand2.GetLength() > 1e-6)
                        hand2 = hand2.Normalize();
                    else
                        hand2 = hand;

                    var handY2 = XYZ.BasisZ.CrossProduct(hand2);
                    handY2 = new XYZ(handY2.X, handY2.Y, 0);
                    if (handY2.GetLength() > 1e-6)
                        handY2 = handY2.Normalize();
                    else
                        handY2 = handY;

                    var dotX = Math.Abs(hand2.DotProduct(mepDir));
                    var dotY = Math.Abs(handY2.DotProduct(mepDir));
                    var expected = wantWidthAlongMep ? dotX : dotY;
                    var alternative = wantWidthAlongMep ? dotY : dotX;

                    if (alternative - expected > 1e-3)
                    {
                        var sign = mepDir.DotProduct(handY2);
                        var fix = sign >= 0 ? (Math.PI / 2.0) : (-Math.PI / 2.0);
                        ElementTransformUtils.RotateElement(_doc, instance.Id, axis, fix);
                        RevitTrace.Warn($"OrientCuboid: axis-swap fix applied instanceId={instance.Id.Value}");
                        
                        // После поворота возвращаем экземпляр в правильную позицию
                        ResetInstanceLocation(instance, instanceLocation);
                    }
                }
                catch
                {
                    // ignore
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("OrientCuboid: orientation fix failed", ex);
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
        /// Вернуть экземпляр в правильную позицию после поворота.
        /// Поворот вокруг оси может сместить экземпляр, если точка вставки семейства
        /// не совпадает с геометрическим центром.
        /// </summary>
        private void ResetInstanceLocation(FamilyInstance instance, XYZ targetLocation)
        {
            try
            {
                // Принудительно обновляем документ, чтобы Location стал доступен
                _doc.Regenerate();
                
                var locationPoint = instance.Location as LocationPoint;
                if (locationPoint != null)
                {
                    var currentLocation = locationPoint.Point;
                    RevitTrace.Info($"ResetInstanceLocation: currentLocation=({currentLocation.X:F4},{currentLocation.Y:F4},{currentLocation.Z:F4}) targetLocation=({targetLocation.X:F4},{targetLocation.Y:F4},{targetLocation.Z:F4})");
                    
                    var offset = targetLocation - currentLocation;
                    
                    if (offset.GetLength() > 1e-6)
                    {
                        // Перемещаем экземпляр обратно в целевую точку
                        ElementTransformUtils.MoveElement(_doc, instance.Id, offset);
                        RevitTrace.Info($"ResetInstanceLocation: moved by offset=({offset.X:F4},{offset.Y:F4},{offset.Z:F4}) length={offset.GetLength() * 304.8:F2}mm");
                    }
                    else
                    {
                        RevitTrace.Info("ResetInstanceLocation: no offset needed");
                    }
                }
                else
                {
                    RevitTrace.Warn("ResetInstanceLocation: LocationPoint is null");
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("ResetInstanceLocation failed", ex);
            }
        }

        /// <summary>
        /// Вернуть экземпляр в правильную позицию по X и Y (Z не трогаем).
        /// Используется после установки параметров, когда геометрия может сместиться.
        /// </summary>
        private void ResetInstanceLocationXY(FamilyInstance instance, XYZ targetLocation)
        {
            try
            {
                var locationPoint = instance.Location as LocationPoint;
                if (locationPoint != null)
                {
                    var currentLocation = locationPoint.Point;
                    
                    // Смещение только по X и Y, Z оставляем как есть
                    var offsetX = targetLocation.X - currentLocation.X;
                    var offsetY = targetLocation.Y - currentLocation.Y;
                    
                    if (Math.Abs(offsetX) > 1e-6 || Math.Abs(offsetY) > 1e-6)
                    {
                        var offset = new XYZ(offsetX, offsetY, 0);
                        ElementTransformUtils.MoveElement(_doc, instance.Id, offset);
                        RevitTrace.Info($"ResetInstanceLocationXY: moved by offset=({offsetX:F4},{offsetY:F4},0) length={offset.GetLength() * 304.8:F2}mm");
                    }
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("ResetInstanceLocationXY failed", ex);
            }
        }

        /// <summary>
        /// Корректировать ориентацию кубика для стены по реальному сечению MEP-элемента.
        /// Для прямоугольных воздуховодов: ширина MEP должна соответствовать ширине кубика,
        /// высота MEP — высоте кубика.
        /// </summary>
        private void OrientCuboidForWallByMepSection(FamilyInstance instance, IntersectionInfo intersection, CuboidParameters cuboidParams, XYZ instanceLocation)
        {
            try
            {
                RevitTrace.Info($"OrientCuboidForWallByMepSection: start instanceId={instance.Id.Value}");

                // Получаем ось Y сечения MEP (направление высоты сечения)
                var sectionYAxis = GetMepSectionYAxis(intersection.MepElement);
                if (sectionYAxis == null)
                {
                    RevitTrace.Warn("OrientCuboidForWallByMepSection: could not get section Y axis");
                    return;
                }

                RevitTrace.Info($"OrientCuboidForWallByMepSection: sectionYAxis=({sectionYAxis.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{sectionYAxis.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{sectionYAxis.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");

                // Для горизонтального MEP, пересекающего стену:
                // - BasisY коннектора (высота сечения) должна быть вертикальной (вдоль Z)
                // - Если она не вертикальная, MEP повёрнут вокруг своей оси

                // Проверяем вертикальную компоненту оси Y сечения
                var sectionYVert = Math.Abs(sectionYAxis.Z);
                
                // Если ось Y сечения почти вертикальная (|Z| > 0.9), ориентация MEP стандартная
                if (sectionYVert > 0.9)
                {
                    RevitTrace.Info("OrientCuboidForWallByMepSection: MEP section Y is vertical, no rotation needed");
                    return;
                }

                // Если ось Y сечения горизонтальная, MEP повёрнут на 90° вокруг своей оси
                // Нужно повернуть кубик на 90° вокруг нормали стены (FacingOrientation)
                var facingDir = instance.FacingOrientation;
                if (facingDir.GetLength() < 1e-6)
                {
                    RevitTrace.Warn("OrientCuboidForWallByMepSection: FacingOrientation is zero");
                    return;
                }

                // Ось вращения - направление "в стену" (Facing), проходящая через точку размещения
                var rotationAxis = Line.CreateBound(
                    instanceLocation,
                    instanceLocation + facingDir);

                // Определяем угол поворота
                // Ось Y сечения должна быть вертикальной (вдоль Z)
                // Текущая ось Y кубика для стены - это вертикаль (Z)
                // Если sectionYAxis горизонтальная, нужно повернуть на 90°

                var sectionYHoriz = new XYZ(sectionYAxis.X, sectionYAxis.Y, 0);
                if (sectionYHoriz.GetLength() < 1e-6)
                {
                    RevitTrace.Info("OrientCuboidForWallByMepSection: sectionY has no horizontal component");
                    return;
                }

                // Поворачиваем кубик на 90° вокруг оси Facing
                ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, Math.PI / 2.0);
                RevitTrace.Info("OrientCuboidForWallByMepSection: rotated 90deg around Facing axis");

                // После поворота нужно также поменять местами Width и Height в параметрах
                // Но это уже сделано в CalculateCuboidParameters на основе MepWidth/MepHeight
            }
            catch (Exception ex)
            {
                RevitTrace.Error("OrientCuboidForWallByMepSection failed", ex);
            }
        }

        /// <summary>
        /// Получить ось Y сечения MEP-элемента из его коннекторов.
        /// Для прямоугольных воздуховодов/лотков это ось вдоль высоты сечения.
        /// </summary>
        private XYZ GetMepSectionYAxis(Element mepElement)
        {
            try
            {
                if (mepElement == null)
                    return null;

                ConnectorManager connMgr = null;

                if (mepElement is MEPCurve mepCurve)
                {
                    connMgr = mepCurve.ConnectorManager;
                }
                else if (mepElement is FamilyInstance fi)
                {
                    var mepModel = fi.MEPModel;
                    if (mepModel != null)
                        connMgr = mepModel.ConnectorManager;
                }

                if (connMgr == null)
                    return null;

                foreach (Connector conn in connMgr.Connectors)
                {
                    if (conn.Shape == ConnectorProfileType.Rectangular ||
                        conn.Shape == ConnectorProfileType.Oval)
                    {
                        var cs = conn.CoordinateSystem;
                        if (cs != null)
                        {
                            RevitTrace.Info($"GetMepSectionYAxis: found connector, BasisY=({cs.BasisY.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{cs.BasisY.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{cs.BasisY.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");
                            return cs.BasisY;
                        }
                    }
                }

                foreach (Connector conn in connMgr.Connectors)
                {
                    var cs = conn.CoordinateSystem;
                    if (cs != null)
                    {
                        return cs.BasisY;
                    }
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("GetMepSectionYAxis failed", ex);
            }

            return null;
        }

        /// <summary>
        /// Получить ось X сечения MEP-элемента из его коннекторов.
        /// Для прямоугольных воздуховодов/лотков это ось вдоль ширины сечения.
        /// </summary>
        private XYZ GetMepSectionXAxis(Element mepElement)
        {
            try
            {
                if (mepElement == null)
                    return null;

                // Получаем ConnectorManager элемента
                ConnectorManager connMgr = null;

                if (mepElement is MEPCurve mepCurve)
                {
                    connMgr = mepCurve.ConnectorManager;
                }
                else if (mepElement is FamilyInstance fi)
                {
                    var mepModel = fi.MEPModel;
                    if (mepModel != null)
                        connMgr = mepModel.ConnectorManager;
                }

                if (connMgr == null)
                    return null;

                // Ищем первый подходящий коннектор с прямоугольным профилем
                foreach (Connector conn in connMgr.Connectors)
                {
                    if (conn.Shape == ConnectorProfileType.Rectangular ||
                        conn.Shape == ConnectorProfileType.Oval)
                    {
                        // CoordinateSystem коннектора: Origin, BasisX, BasisY, BasisZ
                        // BasisX — направление вдоль ширины сечения
                        // BasisY — направление вдоль высоты сечения
                        // BasisZ — направление вдоль оси воздуховода (наружу)
                        var cs = conn.CoordinateSystem;
                        if (cs != null)
                        {
                            RevitTrace.Info($"GetMepSectionXAxis: found connector, BasisX=({cs.BasisX.X.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{cs.BasisX.Y.ToString("G", System.Globalization.CultureInfo.InvariantCulture)},{cs.BasisX.Z.ToString("G", System.Globalization.CultureInfo.InvariantCulture)})");
                            return cs.BasisX;
                        }
                    }
                }

                // Если прямоугольных нет, попробуем любой
                foreach (Connector conn in connMgr.Connectors)
                {
                    var cs = conn.CoordinateSystem;
                    if (cs != null)
                    {
                        return cs.BasisX;
                    }
                }
            }
            catch (Exception ex)
            {
                RevitTrace.Error("GetMepSectionXAxis failed", ex);
            }

            return null;
        }

        /// <summary>
        /// Установить параметры кубика
        /// </summary>
        private void SetCuboidParameters(FamilyInstance instance, CuboidParameters cuboidParams, IntersectionInfo intersection)
        {
            // Ширина - общий параметр для стен и перекрытий
            if (!SetParameterByGuid(instance, CuboidSettings.WidthParamGuid, cuboidParams.Width))
            {
                SetParameterByName(instance, "Ширина", cuboidParams.Width);
            }

            // Второй размер в плоскости зависит от типа элемента вставки:
            // - Для стен: "Высота" (вертикальный размер)
            // - Для перекрытий: "Длина" (второй горизонтальный размер)
            if (intersection.HostType == HostElementType.Wall)
            {
                if (!SetParameterByGuid(instance, CuboidSettings.HeightParamGuid, cuboidParams.Height))
                {
                    SetParameterByName(instance, "Высота", cuboidParams.Height);
                }
            }
            else // Floor
            {
                // Для перекрытий используем параметр "Длина"
                SetParameterByName(instance, "Длина", cuboidParams.Height);
            }

            // Толщина (зависит от типа элемента вставки)
            if (intersection.HostType == HostElementType.Wall)
            {
                if (!SetParameterByGuid(instance, CuboidSettings.WallThicknessParamGuid, intersection.HostThickness))
                {
                    SetParameterByName(instance, "Толщина стены", intersection.HostThickness);
                }
            }
            else
            {
                if (!SetParameterByGuid(instance, CuboidSettings.FloorThicknessParamGuid, intersection.HostThickness))
                {
                    SetParameterByName(instance, "Толщина плиты", intersection.HostThickness);
                }
            }

            // Дополнительная толщина (выступ)
            if (!SetParameterByGuid(instance, CuboidSettings.AdditionalThickness1ParamGuid, cuboidParams.Protrusion))
            {
                SetParameterByName(instance, "Дополнительная толщина 1", cuboidParams.Protrusion);
            }
            if (!SetParameterByGuid(instance, CuboidSettings.AdditionalThickness2ParamGuid, cuboidParams.Protrusion))
            {
                SetParameterByName(instance, "Дополнительная толщина 2", cuboidParams.Protrusion);
            }
        }

        /// <summary>
        /// Установить параметр по GUID
        /// </summary>
        /// <returns>true если параметр успешно установлен</returns>
        private bool SetParameterByGuid(Element element, Guid paramGuid, double value)
        {
            var param = element.get_Parameter(paramGuid);
            if (param != null && !param.IsReadOnly)
            {
                param.Set(value);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Установить параметр по имени
        /// </summary>
        /// <returns>true если параметр успешно установлен</returns>
        private bool SetParameterByName(Element element, string paramName, double value)
        {
            var param = element.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.Double)
            {
                param.Set(value);
                return true;
            }
            return false;
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
