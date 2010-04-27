using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Microsoft.Scripting.Hosting;
using System;
using System.IO;
using System.Collections.Generic;

namespace RevitRubyShell
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    class RunCommand : IExternalCommand
    {
        #region IExternalCommand Members

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RevitRubyShellApplication myapp = RevitRubyShellApplication.GetApplication(commandData);
            ScriptEngine _rubyEngine = myapp.RubyEngine;
            ScriptScope _scope = myapp.RubyScope;            
            _scope.SetVariable("_data", commandData);
            _scope.SetVariable("_app", commandData.Application);

            try
            {
                string source;
                string output = "";
                string commandSrc = myapp.CurrentCommand.Name;
                if(commandSrc == "None")
                    return Result.Failed;
                using (var reader = File.OpenText(commandSrc))
                    source = reader.ReadToEnd();
                var paths = new List<string>();
                foreach (var s in _rubyEngine.GetSearchPaths())
                    paths.Add(s);
                paths.Add(System.IO.Path.GetDirectoryName(commandSrc));
                _rubyEngine.SetSearchPaths(paths);
                if(!myapp.ExecuteCode(source, ref output))
                {
                    message = output;
                    return Result.Failed;
                }

          }
          catch (Exception ex)
          {
              message = ex.Message;
              return Result.Failed;
          }

            return Result.Succeeded;
        }

        #endregion
    }
}
