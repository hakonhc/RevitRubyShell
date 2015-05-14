using System;
using System.Windows;
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
            try
            {
                if (!Install())
                {
                    MessageBox.Show("RevitRubyShell was not installed. No valid Revit 2011-> installation was found", "RevitRubyShell");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "RevitRubyShell");
            }
        }

        private static readonly Guid AppGuid = new Guid("8c90ec3d-b0ef-4b89-be76-96aab8bcd465");
        private const string AppClass = "RevitRubyShell.RevitRubyShellApplication";

        public static bool Install()
        {
            if (RevitProductUtility.GetAllInstalledRevitProducts().Count == 0)
            {
                return false;
            }

            foreach (var product in RevitProductUtility.GetAllInstalledRevitProducts())
            {
                if (product.Version == RevitVersion.Unknown) continue;
                var addinFile = product.CurrentUserAddInFolder + "\\rubyshell.addin";
                var pluginFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + string.Format("\\RevitRubyShell{0}.dll", product.Version);
                var manifest = File.Exists(addinFile) ? AddInManifestUtility.GetRevitAddInManifest(addinFile) : new RevitAddInManifest();

                if (!File.Exists(pluginFile))
                {
                    MessageBox.Show(string.Format("{0} is not supported by this version of RevitRubyShell", product.Version), "RevitRubyShell", MessageBoxButton.OK, MessageBoxImage.Error);
                    continue;
                }

                //Search manifest for app
                RevitAddInApplication app = null;
                foreach (var a in manifest.AddInApplications)
                {
                    if (a.AddInId == AppGuid)
                    {
                        app = a;
                    }
                }

                if (app == null)
                {
                    app = new RevitAddInApplication("RevitRubyShell", pluginFile, AppGuid, AppClass,"NOSYK");
                    manifest.AddInApplications.Add(app);
                }
                else
                {
                    app.Assembly = pluginFile;
                    app.FullClassName = AppClass;
                }

                if (manifest.Name == null)
                {
                    manifest.SaveAs(addinFile);
                }
                else
                {
                    manifest.Save();
                }

                MessageBox.Show(string.Format("RevitRubyShell for {0} was successully installed ", product.Version), "RevitRubyShell");
            }

            return true;
        }
    }
}
