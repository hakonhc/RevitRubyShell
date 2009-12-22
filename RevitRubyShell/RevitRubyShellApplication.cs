using System;
using Autodesk.Revit;

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
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("RevitRubyShell");
            ribbonPanel.AddPushButton("RevitRubyShell", "Open Ruby Shell",
                                      typeof(RevitRubyShellApplication).Assembly.Location,
                                      "RevitRubyShell.ShellCommand");
            return IExternalApplication.Result.Succeeded;
            
        }

        #endregion
    }
}
