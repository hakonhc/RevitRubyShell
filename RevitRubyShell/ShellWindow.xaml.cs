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
using System.Xml.Linq;

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
        public GridSplitter EditorToggle { get { return _editorToggle; } }
        public GridSplitter ConsoleSplitter { get { return _consoleSplitter; } }
        public ToolBar CommandToolbar { get { return cmdToolbar; } }
        #endregion

        private readonly Autodesk.Revit.Application _application;
        private string filename;

        public ShellWindow(ExternalCommandData data)
        {
            InitializeComponent();
            _application = data.Application;

            this.Loaded += (s, e) =>
            {
                var defaultScripts = GetSettings().Root.Descendants("DefaultScript");
                _code.Text = defaultScripts.Count() > 0 ? defaultScripts.First().Value.Replace("\n", "\r\n") : "";
                OutputBuffer = new TextBoxBuffer(_output);

                // Initialize IronRuby
                _rubyEngine = Ruby.CreateEngine();
                _rubyContext = Ruby.GetExecutionContext(_rubyEngine);
                _scope = _rubyEngine.CreateScope();
                _scope.SetVariable("__revit__", _application);
                _scope.SetVariable("_app", _application);
                _scope.SetVariable("_data", data);

                // Cute little trick: warm up the Ruby engine by running some code on another thread:
                new SThread.Thread(new SThread.ThreadStart(() => _rubyEngine.Execute("2 + 2", _scope))).Start();

                // redirect stdout to the output window
                _rubyContext.StandardOutput = OutputBuffer;

                KeyBindings();
                LoadCommands();
            };
        }


        private void LoadCommands()
        {
            foreach (var commandNode in GetSettings().Root.Descendants("Command"))
            {

                var button = new Button();                
                button.Tag = commandNode.Attribute("src").Value;
                button.Content = commandNode.Attribute("name").Value;
                button.Click += CommandClicked;
                CommandToolbar.Items.Add(button);
            }
        }

        void CommandClicked(object sender, EventArgs e)
        {
            string source;
            try
            {
                var commandSrc = (string)(((Button)sender).Tag);
                using (var reader = File.OpenText(commandSrc))
                    source = reader.ReadToEnd();
                var paths = new List<string>();
                foreach (var s in _rubyEngine.GetSearchPaths())
                    paths.Add(s);
                paths.Add(System.IO.Path.GetDirectoryName(commandSrc));
                _rubyEngine.SetSearchPaths(paths);
                ExecuteCode(source);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ex.Message);
            }            
        }

        /// <summary>
        /// Runs all code from a TextBox if there is no selection, otherwise
        /// just runs the selection.
        /// </summary>
        /// <param name="t"></param>
        public void RunCode()
        {
            string code = _code.SelectionLength > 0 ? _code.SelectedText : _code.Text;
            ExecuteCode(code); 
        }

        private void ExecuteCode(string code)
        {
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
            catch (Microsoft.Scripting.SyntaxErrorException e)
            {
                OutputBuffer.write(string.Format("Syntax error at line {1}: {0}\n", e.Message, e.Line));
            }
            catch (Exception e)
            {
                var exceptionService = _rubyEngine.GetService<ExceptionOperations>();
                string message, typeName;
                exceptionService.GetExceptionMessage(e, out message, out typeName);
                OutputBuffer.write(string.Format("{0} ({1})\n", message, typeName));
            }
        }


        /// <summary>
        /// When Ctrl-Enter is pressed, run the script code
        /// </summary>
        private void KeyBindings()
        {
            _code.KeyDown += (se, args) =>
            {
                if(args.IsCtrl(Key.Enter))
                    RunCode();
                else if(args.IsCtrl(Key.W))
                    Output.Clear();
                else if(args.IsCtrl(Key.S))
                    save_code(null, null);

            };
        }

        //Save code buffer
        private void save_code(object sender, RoutedEventArgs e)
        {
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = filename == null ? "command.rb" : filename; // Default file name
            dlg.DefaultExt = ".rb"; // Default file extension
            dlg.Filter = "Ruby code (.rb)|*.rb"; // Filter files by extension

            // Process save file dialog box results
            if (dlg.ShowDialog() == true)
            {
                // Save document
                this.filename = dlg.FileName;
                File.WriteAllText(this.filename, _code.Text);
                this.Title = filename;
            }

        }

        //Open rb file
        private void open_code(object sender, RoutedEventArgs e)
        {

            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = filename == null ? "command.rb" : filename; // Default file name
            dlg.DefaultExt = ".rb"; // Default file extension
            dlg.Filter = "Ruby code (.rb)|*.rb"; // Filter files by extension

            // Process open file dialog box results
            if (dlg.ShowDialog() == true)
            {
                // Open document
                _code.Text = File.ReadAllText(dlg.FileName);
            }

        }

        private static XDocument GetSettings()
        {
            //Whould be nice to use YAML instead!
            string assemblyFolder = new FileInfo(typeof(RevitRubyShellApplication).Assembly.Location).DirectoryName;
            string settingsFile = System.IO.Path.Combine(assemblyFolder, "RevitRubyShell.xml");
            return XDocument.Load(settingsFile);
        }

        private void run_code(object sender, RoutedEventArgs e)
        {
            RunCode();
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

    #region Extension Methods

    public static class ExtensionMethods
    {
        public static bool IsCtrl(this KeyEventArgs keyEvent, Key value)
        {
            return keyEvent.KeyboardDevice.Modifiers == ModifierKeys.Control
                && keyEvent.Key == value;
        }

        public static bool IsCtrlShift(this KeyEventArgs keyEvent, Key value)
        {
            return keyEvent.KeyboardDevice.Modifiers == ModifierKeys.Control
                && keyEvent.KeyboardDevice.Modifiers == ModifierKeys.Shift
                && keyEvent.Key == value;
        }

        public static bool Is(this KeyEventArgs keyEvent, Key value)
        {
            return keyEvent.KeyboardDevice.Modifiers == ModifierKeys.None
                && keyEvent.Key == value;
        }
    }

    #endregion
}