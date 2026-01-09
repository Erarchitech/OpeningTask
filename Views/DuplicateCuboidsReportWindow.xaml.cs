using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace OpeningTask.Views
{
    public partial class DuplicateCuboidsReportWindow : Window
    {
        private readonly string _reportText;

        public DuplicateCuboidsReportWindow(IEnumerable<long> duplicateIds)
        {
            InitializeComponent();

            var ids = (duplicateIds ?? Enumerable.Empty<long>())
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            _reportText = BuildReport(ids);

            IdsItemsControl.ItemsSource = ids.Select(x => x.ToString()).ToList();
            ReportTextBox.Text = _reportText;

            OkButton.Click += OkButton_Click;
            SaveButton.Click += SaveButton_Click;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new SaveFileDialog
                {
                    Filter = "Текстовые файлы (*.txt)|*.txt",
                    DefaultExt = ".txt",
                    FileName = "ОтчётДубликатыБоксов.txt"
                };

                if (dialog.ShowDialog(this) == true)
                {
                    File.WriteAllText(dialog.FileName, _reportText, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сохранить файл: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string BuildReport(IReadOnlyList<long> ids)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Отчёт о дубликатах боксов");
            sb.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine();
            sb.AppendLine("ID существующих боксов (обнаружены дубликаты):");

            if (ids == null || ids.Count == 0)
            {
                sb.AppendLine("(нет)");
                return sb.ToString();
            }   

            foreach (var id in ids)
            {
                sb.AppendLine(id.ToString());
            }   

            return sb.ToString();
        }
    }
}