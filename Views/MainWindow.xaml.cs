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

            // Подписка на изменения видимости для обновления привязок
            this.IsVisibleChanged += MainWindow_IsVisibleChanged;
        }

        private void MainWindow_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true)
            {
                // Принудительное обновление привязок ItemsControl при появлении окна
                MepModelsControl?.Items.Refresh();
                ArKrModelsControl?.Items.Refresh();
            }
        }
    }
}
