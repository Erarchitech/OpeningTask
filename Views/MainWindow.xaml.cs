using OpeningTask.ViewModels;
using System.Windows;

namespace OpeningTask.Views
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainWindowViewModel _viewModel;

        public MainWindow(MainWindowViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;
            viewModel.SetWindow(this);

            // Subscribe to visibility changes to refresh bindings
            this.IsVisibleChanged += MainWindow_IsVisibleChanged;
        }

        private void MainWindow_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true)
            {
                // Force refresh of ItemsControl bindings when window becomes visible
                MepModelsControl?.Items.Refresh();
                ArKrModelsControl?.Items.Refresh();
            }
        }
    }
}
