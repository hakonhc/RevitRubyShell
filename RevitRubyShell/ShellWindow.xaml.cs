using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Timers;
using System.IO;
using Autodesk.Revit;
using System.Diagnostics;
using SThread = System.Threading;
using Microsoft.Scripting.Hosting;
using IronRuby;

namespace RevitRubyShell
{

    public partial class ShellWindow : Window
    {

        #region Running code
        public TextBoxBuffer OutputBuffer { get; internal set; }
        private bool _isCtrlPressed;
        private ScriptEngine _rubyEngine;
        public ScriptEngine RubyEngine { get { return _rubyEngine; } }
        private IronRuby.Runtime.RubyContext _rubyContext;
        private ScriptScope _scope;
        #endregion

        #region UI accessors
        public TextBox History { get { return _history; } }
        public TextBox Output { get { return _output; } }
        public TextBox Code { get { return _code; } }
        public StackPanel CanvasControls { get { return _canvasControls; } }
        public StackPanel OutputControls { get { return _outputControls; } }
        public GridSplitter EditorToggle { get { return _editorToggle; } }
        public GridSplitter ConsoleSplitter { get { return _consoleSplitter; } }
        #endregion

        private readonly Autodesk.Revit.Application _application;

        public ShellWindow(Autodesk.Revit.Application app)
        {
            InitializeComponent();
            _application = app;

            this.Loaded += (s, e) =>
            {
                #region Wecome Text
                _code.Text = @"# Welcome to Ruby Shell!

# All Ruby code typed here can be run by pressing
# Ctrl-Enter. If you don't want to run everything,
# just select the text you wan to run and press
# the same key combination. Ctrl-C will empty the output.

# You can use the special __revit__ variable to get hold of the Application object

# This is nothing more than a Ruby interpreter:
# Try the following; it will print to the output
# window below the code:

10.times{|i| puts i * i}
";
                #endregion

                OutputBuffer = new TextBoxBuffer(_output);

                // Initialize IronRuby

                _rubyEngine = Ruby.CreateEngine();
                _rubyContext = Ruby.GetExecutionContext(_rubyEngine);
                _scope = _rubyEngine.CreateScope();
                _scope.SetVariable("__revit__", _application);

                // Cute little trick: warm up the Ruby engine by running some code on another thread:
                new SThread.Thread(new SThread.ThreadStart(() => _rubyEngine.Execute("2 + 2", _scope))).Start();

                // redirect stdout to the output window
                _rubyContext.StandardOutput = OutputBuffer;

                KeyBindings();
            };
        }

        /// <summary>
        /// Runs all code from a TextBox if there is no selection, otherwise
        /// just runs the selection.
        /// </summary>
        /// <param name="t"></param>
        public void RunCode(TextBox t)
        {
            string code = t.SelectionLength > 0 ? t.SelectedText : t.Text;
            try
            {
                // Run the code
                var result = _rubyEngine.Execute(code, _scope);
             
                // write the result to the output window
                var output = string.Format("=> {0}\n", _rubyContext.Inspect(result));
                OutputBuffer.write(output);

                // add the code to the history
                _history.AppendText(string.Format("{0}\n# {1}", code, output));

            }
            catch (Exception e)
            {
                OutputBuffer.write(string.Format("ERROR => {0}\n",e.Message));
            }
        }
    
        /// <summary>
        /// When Ctrl-Enter is pressed, run the script code
        /// </summary>
        private void KeyBindings()
        {
            _code.KeyDown += (se, args) =>
            {
                if (args.Key == Key.LeftCtrl || args.Key == Key.RightCtrl)
                    _isCtrlPressed = true;
                if (_isCtrlPressed && args.Key == Key.Enter)
                    RunCode(_code);
                if (_isCtrlPressed && args.Key == Key.C)
                    Output.Clear();                   
                    
            };
            _code.KeyUp += (se, args) =>
            {
                if (args.Key == Key.LeftCtrl || args.Key == Key.RightCtrl)
                    _isCtrlPressed = false;
            };
        }    
    }

    /// <summary>
    /// Simple TextBox Buffer class
    /// </summary>
    public class TextBoxBuffer
    {
        private TextBox box;

        public TextBoxBuffer(TextBox t)
        {
            box = t;
        }

        public void write(string str)
        {
            box.Dispatcher.BeginInvoke((Action)(() =>
            {
                box.AppendText(str);
                box.ScrollToEnd();
            }));
        }
    }
}