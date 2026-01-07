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
            string ribbonName = "MsMind";
            try
            {
                application.CreateRibbonTab(ribbonName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException) { }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData button = new PushButtonData("OpeningTask", "Задания\nна отверстия", assemblyPath, "OpeningTask.Commands.OpeningTaskCommand");
            var panel = application.CreateRibbonPanel(ribbonName, "Task");
            panel.AddItem(button);

            return Result.Succeeded;
        }
    }
}
