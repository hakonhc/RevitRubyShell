using System;
using Autodesk.Revit.UI;
using System.Xml.Linq;
using System.IO;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;

namespace RevitRubyShell
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class RevitRubyShellApplication : IExternalApplication
    {
        #region IExternalApplication Members

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Ruby scripting");
            PushButton pushButton = ribbonPanel.AddItem(new PushButtonData("RevitRubyShell", "Open Shell",
                                       typeof(RevitRubyShellApplication).Assembly.Location,
                                      "RevitRubyShell.ShellCommand")) as PushButton;
            pushButton.LargeImage = GetImage("console-5.png");
            return Result.Succeeded;
            
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
