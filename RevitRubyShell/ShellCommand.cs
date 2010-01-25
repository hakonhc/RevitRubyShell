using Autodesk.Revit;

namespace RevitRubyShell
{
    public class ShellCommand : IExternalCommand
    {    
        public IExternalCommand.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var gui = new ShellWindow(commandData);
            gui.Show();
            gui.BringIntoView();
            return IExternalCommand.Result.Succeeded;
        }
    }
}
