using Autodesk.Revit.DB;
using System;

namespace OpeningTask.Models
{
    /// <summary>
    /// Информация о связанной модели Revit
    /// </summary>
    public class LinkedModelInfo : ViewModels.BaseViewModel
    {
        private bool _isSelected;
        private string _name;
        private RevitLinkInstance _linkInstance;
        private Document _linkedDocument;

        public LinkedModelInfo(RevitLinkInstance linkInstance)
        {
            _linkInstance = linkInstance ?? throw new ArgumentNullException(nameof(linkInstance));
            _linkedDocument = linkInstance.GetLinkDocument();
            _name = _linkedDocument?.Title ?? linkInstance.Name;
        }

        /// <summary>
        /// Выбрана ли модель
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        /// <summary>
        /// Имя связанной модели
        /// </summary>
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        /// <summary>
        /// Экземпляр связи Revit
        /// </summary>
        public RevitLinkInstance LinkInstance => _linkInstance;

        /// <summary>
        /// Документ связанной модели
        /// </summary>
        public Document LinkedDocument => _linkedDocument;

        /// <summary>
        /// Проверка загружена ли связь
        /// </summary>
        public bool IsLoaded => _linkedDocument != null;

        /// <summary>
        /// Трансформация связанной модели
        /// </summary>
        public Transform Transform => _linkInstance.GetTotalTransform();
    }
}
