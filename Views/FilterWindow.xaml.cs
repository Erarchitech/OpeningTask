using OpeningTask.ViewModels;
using System.Windows;

namespace OpeningTask.Views
{
    /// <summary>
    /// Логика взаимодействия для FilterWindow.xaml
    /// </summary>
    public partial class FilterWindow : Window
    {
        public FilterWindow(FilterWindowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            viewModel.SetWindow(this);
        }
    }
}
