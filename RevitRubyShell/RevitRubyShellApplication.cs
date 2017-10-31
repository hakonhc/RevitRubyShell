using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using IronRuby;
using Microsoft.Scripting.Hosting;
using SThread = System.Threading;

namespace RevitRubyShell
{
    [Regeneration(RegenerationOption.Manual)]
    class RevitRubyShellApplication : IExternalApplication
    {
        public static RevitRubyShellApplication RevitRubyShell;

        //Ironruby
        private ScriptEngine _rubyEngine;
        private ScriptScope _scope;
    
        public ScriptEngine RubyEngine { get { return _rubyEngine; } }        
        public ScriptScope RubyScope { get { return _scope; } }

        public static Guid DockGuid = new Guid("{C584FF49-4869-4704-BB18-820F7F13F640}");

        public Queue<Action<UIApplication>> Queue = new Queue<Action<UIApplication>>(); 
      
        #region IExternalApplication Members

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            //Create panel
            var ribbonPanel = application.CreateRibbonPanel("Ruby scripting");
            var pushButton = ribbonPanel.AddItem(
                new PushButtonData(
                    "RevitRubyShell", 
                    "Open Shell",
                    typeof(RevitRubyShellApplication).Assembly.Location,
                    "RevitRubyShell.ShellCommand")) as PushButton;

            pushButton.LargeImage = GetImage("console-5.png");
            
            //Start ruby interpreter
            _rubyEngine = Ruby.CreateEngine();
            _scope = _rubyEngine.CreateScope();

            // Warm up the Ruby engine by running some code on another thread:
            new SThread.Thread(
                () => 
                {
                    var defaultScripts = GetSettings().Root.Descendants("OnLoad").ToArray();
                    var script = defaultScripts.Any() ? defaultScripts.First().Value.Replace("\n", "\r\n") : "";
                    _rubyEngine.Execute(script, _scope);
                } ).Start();

            RevitRubyShellApplication.RevitRubyShell = this;

            application.Idling += (sender, args) =>
                {
                    var uiapp = sender as UIApplication; 
                    lock (this.Queue)
                    {
                        if (this.Queue.Count <= 0) return;

                        var task = this.Queue.Dequeue();

                        // execute the task!
                        try
                        {
                            task(uiapp);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                        }
                    }
                };

            var win = new ShellWindow();       
            application.RegisterDockablePane(new DockablePaneId(DockGuid), "RevitRubyShell", win);
        
            return Result.Succeeded;
        }
        
        #endregion

        #region App Icon handling
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
   
        public static XDocument GetSettings()
        {
            // Whould be nice to use YAML instead!
            var assemblyFolder = new FileInfo(typeof(RevitRubyShellApplication).Assembly.Location).DirectoryName;
            var settingsFile = Path.Combine(assemblyFolder, "RevitRubyShell.xml");
            return XDocument.Load(settingsFile);
        }

        public string LastCode { get; set; }
  
        public bool ExecuteCode(string code, out string output)
        {
            try
            {
                // Run the code
                var result = _rubyEngine.Execute(code, _scope);
                // Write the result to the output window
                output = string.Format("=> {0}\n", ((IronRuby.Runtime.RubyContext)Microsoft.Scripting.Hosting.Providers.HostingHelpers.GetLanguageContext(_rubyEngine)).Inspect(result));
                return true;              
            }
            catch (Microsoft.Scripting.SyntaxErrorException e)
            {
                output = string.Format("Syntax error at line {1}: {0}\n", e.Message, e.Line);
            }
            catch (Exception e)
            {
                var exceptionService = _rubyEngine.GetService<ExceptionOperations>();
                string message, typeName;
                exceptionService.GetExceptionMessage(e, out message, out typeName);
                output = string.Format("{0} ({1})\n", message, typeName);                
            }

            return false;
        }
    }
}
