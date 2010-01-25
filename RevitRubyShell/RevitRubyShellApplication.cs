using System;
using Autodesk.Revit;
using System.Xml.Linq;
using System.IO;
using System.Windows.Media.Imaging;

namespace RevitRubyShell
{
    class RevitRubyShellApplication : IExternalApplication
    {
        #region IExternalApplication Members

        public IExternalApplication.Result OnShutdown(ControlledApplication application)
        {
            return IExternalApplication.Result.Succeeded;
        }

        public IExternalApplication.Result OnStartup(ControlledApplication application)
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Ruby scripting");
            var pb = ribbonPanel.AddPushButton("RevitRubyShell", "Open Shell",
                                       typeof(RevitRubyShellApplication).Assembly.Location,
                                      "RevitRubyShell.ShellCommand");
            pb.LargeImage = new BitmapImage(new Uri(@"console.ico"));
            return IExternalApplication.Result.Succeeded;
            
        }
     #endregion
    }
}
