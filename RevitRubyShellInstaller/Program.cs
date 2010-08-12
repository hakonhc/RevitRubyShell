using System;
using Autodesk.RevitAddIns;
using System.IO;
using System.Reflection;

namespace RevitRubyShellInstaller
{
    /// <summary>
    /// Installes the RevitRubyShell in Revit with a manifest file
    /// </summary>
    class RevitRubyShellInstaller
    {
        static void Main(string[] args)
        {
            if(install())
                System.Windows.MessageBox.Show("RevitRubyShell was successully installed ", "RevitRubyShell");
            else
                System.Windows.MessageBox.Show("RevitRubyShell was not installed. No valid Revit 2011-> installation was found", "RevitRubyShell");
        }

        private static Guid APP_GUID = new Guid("8c90ec3d-b0ef-4b89-be76-96aab8bcd465");
        private const string APP_CLASS = "RevitRubyShell.RevitRubyShellApplication";

        public static bool install()
        {
            if (RevitProductUtility.GetAllInstalledRevitProducts().Count == 0)
                return false;
            foreach (RevitProduct product in RevitProductUtility.GetAllInstalledRevitProducts())
            {
                string addinFile = product.CurrentUserAddInFolder + "\\rubyshell.addin";
                string pluginFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\RevitRubyShell.dll";

                RevitAddInManifest manifest;
                if (File.Exists(addinFile))
                    manifest = AddInManifestUtility.GetRevitAddInManifest(addinFile);
                else
                    manifest = new RevitAddInManifest();

                //Search manifest for app
                RevitAddInApplication app = null;
                foreach (RevitAddInApplication a in manifest.AddInApplications)
                {
                    if (a.AddInId == APP_GUID)
                        app = a;
                }
                if (app == null)
                {
                    app = new RevitAddInApplication("RevitRubyShell", pluginFile, APP_GUID, APP_CLASS);
                    manifest.AddInApplications.Add(app);
                }
                else
                {
                    app.Assembly = pluginFile;
                    app.FullClassName = APP_CLASS;
                }

                if (manifest.Name == null)
                    manifest.SaveAs(addinFile);
                else
                    manifest.Save();
            }
            return true;
        }
    }
}
