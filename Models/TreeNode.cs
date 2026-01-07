using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace OpeningTask.Models
{
    /// <summary>
    /// Узел дерева для отображения типов и параметров
    /// </summary>
    public class TreeNode : ViewModels.BaseViewModel
    {
        private bool? _isChecked = false;
        private string _name;
        private TreeNode _parent;
        private ObservableCollection<TreeNode> _children;
        private object _tag;
        private bool _isExpanded;

        public TreeNode(string name, TreeNode parent = null, object tag = null)
        {
            _name = name;
            _parent = parent;
            _tag = tag;
            _children = new ObservableCollection<TreeNode>();
        }

        /// <summary>
        /// Имя узла
        /// </summary>
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        /// <summary>
        /// Состояние выбора (null = неопределенное, только для родительских узлов)
        /// </summary>
        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                // Для листовых узлов (детей) разрешены только значения true/false
                var newValue = value;
                if (IsLeaf && !newValue.HasValue)
                {
                    newValue = false;
                }

                if (SetProperty(ref _isChecked, newValue))
                {
                    // Если значение определенное, обновляем дочерние элементы
                    if (newValue.HasValue && _children.Count > 0)
                    {
                        UpdateChildren(newValue.Value);
                    }

                    // Обновляем родительский элемент
                    UpdateParent();
                }
            }
        }

        /// <summary>
        /// Является ли узел листовым (без детей, уровень ниже родителя)
        /// </summary>
        public bool IsLeaf => _parent != null && _children.Count == 0;

        /// <summary>
        /// Поддерживает ли узел трехсторонний чекбокс (только для родительских узлов)
        /// </summary>
        public bool IsThreeState => !IsLeaf;

        /// <summary>
        /// Развернут ли узел
        /// </summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set => SetProperty(ref _isExpanded, value);
        }

        /// <summary>
        /// Родительский узел
        /// </summary>
        public TreeNode Parent
        {
            get => _parent;
            set => SetProperty(ref _parent, value);
        }

        /// <summary>
        /// Дочерние узлы
        /// </summary>
        public ObservableCollection<TreeNode> Children
        {
            get => _children;
            set => SetProperty(ref _children, value);
        }

        /// <summary>
        /// Дополнительные данные узла
        /// </summary>
        public object Tag
        {
            get => _tag;
            set => SetProperty(ref _tag, value);
        }

        /// <summary>
        /// Обновление состояния дочерних элементов
        /// </summary>
        private void UpdateChildren(bool isChecked)
        {
            foreach (var child in _children)
            {
                child._isChecked = isChecked;
                child.OnPropertyChanged(nameof(IsChecked));
                child.UpdateChildren(isChecked);
            }
        }

        /// <summary>
        /// Обновление состояния родительского элемента
        /// </summary>
        private void UpdateParent()
        {
            if (_parent == null) return;

            var checkedCount = _parent._children.Count(c => c.IsChecked == true);
            var uncheckedCount = _parent._children.Count(c => c.IsChecked == false);

            if (checkedCount == _parent._children.Count)
            {
                _parent._isChecked = true;
            }
            else if (uncheckedCount == _parent._children.Count)
            {
                _parent._isChecked = false;
            }
            else
            {
                _parent._isChecked = null; // Неопределенное состояние
            }

            _parent.OnPropertyChanged(nameof(IsChecked));
            _parent.UpdateParent();
        }

        /// <summary>
        /// Добавление дочернего узла
        /// </summary>
        public TreeNode AddChild(string name, object tag = null)
        {
            var child = new TreeNode(name, this, tag);
            _children.Add(child);
            OnPropertyChanged(nameof(IsLeaf));
            OnPropertyChanged(nameof(IsThreeState));
            return child;
        }

        /// <summary>
        /// Установка состояния без обновления иерархии
        /// </summary>
        public void SetCheckedSilent(bool? value)
        {
            // Для листовых узлов разрешены только значения true/false
            if (IsLeaf && !value.HasValue)
            {
                value = false;
            }
            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));
        }
    }
}
