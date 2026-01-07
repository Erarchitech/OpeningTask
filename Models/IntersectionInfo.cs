using Autodesk.Revit.DB;
using System;

namespace OpeningTask.Models
{
    /// <summary>
    /// Тип сечения элемента инженерной системы
    /// </summary>
    public enum MepSectionType
    {
        /// <summary>
        /// Круглое сечение
        /// </summary>
        Round,

        /// <summary>
        /// Прямоугольное сечение
        /// </summary>
        Rectangular
    }

    /// <summary>
    /// Тип элемента инженерной системы
    /// </summary>
    public enum MepElementType
    {
        /// <summary>
        /// Труба
        /// </summary>
        Pipe,

        /// <summary>
        /// Воздуховод
        /// </summary>
        Duct,

        /// <summary>
        /// Кабельный лоток / короб
        /// </summary>
        Tray,

        /// <summary>
        /// Неизвестный тип
        /// </summary>
        Unknown
    }

    /// <summary>
    /// Тип элемента вставки (стена/перекрытие)
    /// </summary>
    public enum HostElementType
    {
        /// <summary>
        /// Стена
        /// </summary>
        Wall,

        /// <summary>
        /// Перекрытие
        /// </summary>
        Floor
    }

    /// <summary>
    /// Информация о пересечении элемента инженерной системы со стеной/перекрытием
    /// </summary>
    public class IntersectionInfo
    {
        /// <summary>
        /// Элемент инженерной системы (из связанной модели)
        /// </summary>
        public Element MepElement { get; set; }

        /// <summary>
        /// Связанная модель с MEP элементом
        /// </summary>
        public RevitLinkInstance MepLinkInstance { get; set; }

        /// <summary>
        /// Элемент вставки (стена/перекрытие из связанной модели)
        /// </summary>
        public Element HostElement { get; set; }

        /// <summary>
        /// Связанная модель с элементом вставки
        /// </summary>
        public RevitLinkInstance HostLinkInstance { get; set; }

        /// <summary>
        /// Точка вставки кубика (в координатах текущего документа)
        /// </summary>
        public XYZ InsertionPoint { get; set; }

        /// <summary>
        /// Направление (нормаль) элемента вставки
        /// </summary>
        public XYZ HostNormal { get; set; }

        /// <summary>
        /// Направление элемента инженерной системы
        /// </summary>
        public XYZ MepDirection { get; set; }

        /// <summary>
        /// Тип элемента инженерной системы
        /// </summary>
        public MepElementType MepType { get; set; }

        /// <summary>
        /// Тип сечения элемента
        /// </summary>
        public MepSectionType SectionType { get; set; }

        /// <summary>
        /// Тип элемента вставки
        /// </summary>
        public HostElementType HostType { get; set; }

        /// <summary>
        /// Ширина сечения MEP элемента (футы)
        /// </summary>
        public double MepWidth { get; set; }

        /// <summary>
        /// Высота сечения MEP элемента (футы)
        /// </summary>
        public double MepHeight { get; set; }

        /// <summary>
        /// Диаметр MEP элемента для круглого сечения (футы)
        /// </summary>
        public double MepDiameter { get; set; }

        /// <summary>
        /// Толщина элемента вставки (футы)
        /// </summary>
        public double HostThickness { get; set; }
    }
}