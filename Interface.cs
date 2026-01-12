using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OpeningTask
{
    public class Interface : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string ribbonName = "TASK";
            try
            {
                application.CreateRibbonTab(ribbonName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException) { }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData button = new PushButtonData("OpeningTask", "Задания\nна отверстия", assemblyPath, "OpeningTask.Commands.OpeningTaskCommand");
            
            // Иконка для кнопки (32x32 для LargeImage)
            try
            {
                var largeIcon = new System.Windows.Media.Imaging.BitmapImage();
                largeIcon.BeginInit();
                largeIcon.UriSource = new Uri("pack://application:,,,/OpeningTask;component/Resources/Images/OpeningTaskButton.png");
                largeIcon.DecodePixelWidth = 32;
                largeIcon.DecodePixelHeight = 32;
                largeIcon.EndInit();
                button.LargeImage = largeIcon;
                
                // Маленькая иконка (16x16)
                var smallIcon = new System.Windows.Media.Imaging.BitmapImage();
                smallIcon.BeginInit();
                smallIcon.UriSource = new Uri("pack://application:,,,/OpeningTask;component/Resources/Images/OpeningTaskButton.png");
                smallIcon.DecodePixelWidth = 16;
                smallIcon.DecodePixelHeight = 16;
                smallIcon.EndInit();
                button.Image = smallIcon;
            }
            catch { }
            
            var panel = application.CreateRibbonPanel(ribbonName, "Task");
            panel.AddItem(button);

            return Result.Succeeded;
        }
    }
}
