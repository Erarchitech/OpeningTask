using System;

namespace OpeningTask.Models
{
    /// <summary>
    /// Настройки для создания кубиков
    /// </summary>
    public class CuboidSettings : ViewModels.BaseViewModel
    {
        private int _roundingValue = 50;
        private int _minOffset = 30;
        private int _protrusion = 100;
        private bool _usePipeRoundCuboid = true;
        private bool _useDuctRoundCuboid = false;

        /// <summary>
        /// Значение округления размеров (мм)
        /// </summary>
        public int RoundingValue
        {
            get => _roundingValue;
            set => SetProperty(ref _roundingValue, value);
        }

        /// <summary>
        /// Минимальный отступ от элемента (мм)
        /// </summary>
        public int MinOffset
        {
            get => _minOffset;
            set => SetProperty(ref _minOffset, value);
        }

        /// <summary>
        /// Выступ от элемента вставки с каждой стороны (мм)
        /// </summary>
        public int Protrusion
        {
            get => _protrusion;
            set => SetProperty(ref _protrusion, value);
        }

        /// <summary>
        /// Использовать круглый кубик для труб круглого сечения
        /// </summary>
        public bool UsePipeRoundCuboid
        {
            get => _usePipeRoundCuboid;
            set => SetProperty(ref _usePipeRoundCuboid, value);
        }

        /// <summary>
        /// Использовать круглый кубик для воздуховодов круглого сечения
        /// </summary>
        public bool UseDuctRoundCuboid
        {
            get => _useDuctRoundCuboid;
            set => SetProperty(ref _useDuctRoundCuboid, value);
        }

        /// <summary>
        /// Путь к папке с семействами кубиков
        /// </summary>
        public string FamilyFolderPath => @"C:\OpeningTaskResources";

        /// <summary>
        /// GUID параметра "Ширина"
        /// </summary>
        public static readonly Guid WidthParamGuid = new Guid("6f459bf2-cf72-4223-9ee8-78e8252046a0");

        /// <summary>
        /// GUID параметра "Высота"
        /// </summary>
        public static readonly Guid HeightParamGuid = new Guid("60bf9b18-17f9-4b8f-b214-fc13bc7b357f");

        /// <summary>
        /// GUID параметра "Толщина стены"
        /// </summary>
        public static readonly Guid WallThicknessParamGuid = new Guid("6df7db81-c1d3-48f5-97a1-cd35960d9f1c");

        /// <summary>
        /// GUID параметра "Толщина плиты"
        /// </summary>
        public static readonly Guid FloorThicknessParamGuid = new Guid("6b790a90-bd86-4366-84ee-8a6e60af2288");

        /// <summary>
        /// GUID параметра "Дополнительная толщина 1"
        /// </summary>
        public static readonly Guid AdditionalThickness1ParamGuid = new Guid("ed4c28e9-f16d-49e8-bb98-7be58cdc2893");

        /// <summary>
        /// GUID параметра "Дополнительная толщина 2"
        /// </summary>
        public static readonly Guid AdditionalThickness2ParamGuid = new Guid("1362f685-6d3d-4b3c-8a6d-51c59e1fd44b");

        /// <summary>
        /// Получить имя файла семейства кубика
        /// </summary>
        public string GetCuboidFamilyName(HostElementType hostType, MepSectionType sectionType, MepElementType mepType)
        {
            bool useRoundCuboid = false;

            if (sectionType == MepSectionType.Round)
            {
                if (mepType == MepElementType.Pipe && UsePipeRoundCuboid)
                    useRoundCuboid = true;
                else if (mepType == MepElementType.Duct && UseDuctRoundCuboid)
                    useRoundCuboid = true;
            }

            string hostPart = hostType == HostElementType.Wall ? "Стена" : "Перекрытие";
            string shapePart = useRoundCuboid ? "Круг" : "Прямоугольный";

            return $"Кубик_{hostPart}_{shapePart}";
        }

        /// <summary>
        /// Получить полный путь к файлу семейства
        /// </summary>
        public string GetCuboidFamilyPath(HostElementType hostType, MepSectionType sectionType, MepElementType mepType)
        {
            var familyName = GetCuboidFamilyName(hostType, sectionType, mepType);
            return System.IO.Path.Combine(FamilyFolderPath, familyName + ".rfa");
        }
    }
}