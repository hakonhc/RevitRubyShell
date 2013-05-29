using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;

namespace RevitRubyShell
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class ShellCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var dp = commandData.Application.GetDockablePane(new DockablePaneId(RevitRubyShellApplication.DockGuid));
            dp.Show();
            return Result.Succeeded;
        }
    } 
}
