using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace RevitRubyShell
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class ShellCommand : IExternalCommand
    {
        private ShellWindow win;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            win = new ShellWindow(commandData);
            win.ShowDialog();
            return Result.Succeeded;
        }
    } 
}
