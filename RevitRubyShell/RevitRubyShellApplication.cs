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
            pb.LargeImage = GetImage("console-5.png");
            return IExternalApplication.Result.Succeeded;
            
        }
        private BitmapImage GetImage(string resourcePath)
        {
            var image = new BitmapImage();

            string moduleName = this.GetType().Assembly.GetName().Name;
            string resourceLocation =
                string.Format("pack://application:,,,/{0};component/{1}", moduleName,
                              resourcePath);

            try
            {
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                image.UriSource = new Uri(resourceLocation);
                image.EndInit();
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.ToString());
            }

            return image;
        }
     #endregion
    }
}
