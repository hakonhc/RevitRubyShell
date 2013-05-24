using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Shapes;
using Microsoft.Scripting.Hosting;

namespace RevitRubyShell
{
    public partial class ShellWindow : Window
    {
        #region Running code
        public TextBoxBuffer OutputBuffer { get; internal set; }
        private ScriptEngine _rubyEngine;
        private ScriptScope _scope;
        private IronRuby.Runtime.RubyContext _rubyContext;
        #endregion

        #region UI accessors
        public TextBox History { get { return _history; } }
        public TextBox Output { get { return _output; } }
        public RichTextBox Code { get { return _code; } }
        public GridSplitter EditorToggle { get { return _editorToggle; } }
        public GridSplitter ConsoleSplitter { get { return _consoleSplitter; } }
        #endregion

        private readonly Autodesk.Revit.UI.UIApplication _application;
        private string filename;
        private RevitRubyShellApplication myapp;

        public ShellWindow(Autodesk.Revit.UI.ExternalCommandData data)
        {
            InitializeComponent();
            _application = data.Application;
            myapp = RevitRubyShellApplication.GetApplication(data); 

            this.Loaded += (s, e) =>
                {                    
                    var defaultScripts = RevitRubyShellApplication.GetSettings().Root.Descendants("DefaultScript");
                    var lastCode = myapp.LastCode;
                    if (string.IsNullOrEmpty(lastCode))
                    {
                        _code.SetText(defaultScripts.Count() > 0 ? defaultScripts.First().Value.Replace("\n", "\r\n") : "");
                    }
                    else
                    {
                        _code.SetText(lastCode);
                    }

                    OutputBuffer = new TextBoxBuffer(_output);

                    // Initialize IronRuby
                    _rubyEngine = myapp.RubyEngine;
                    _scope = myapp.RubyScope;
                    _rubyContext = (IronRuby.Runtime.RubyContext)Microsoft.Scripting.Hosting.Providers.HostingHelpers.GetLanguageContext(_rubyEngine);          
                    _scope.SetVariable("_data", data);
                    _scope.SetVariable("_app", data.Application);

                    // redirect stdout to the output window
                    _rubyContext.StandardOutput = OutputBuffer;                
                    KeyBindings();
                };

            this.Closed += (s, e) =>
                {
                    myapp.LastCode = _code.GetText();
                };
        }


        /// <summary>
        /// Runs all code from a TextBox if there is no selection, otherwise
        /// just runs the selection.
        /// </summary>
        /// <param name="t"></param>
        public void RunCode()
        {
            string code = _code.GetText();
            string output = "";
      
            bool result = myapp.ExecuteCode(code, ref output);
            if (result)
            {
                OutputBuffer.write(output);
                // add the code to the history
                _history.AppendText(string.Format("{0}\n# {1}", code, output));
            }
            else
            {
                OutputBuffer.write(output);
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
                File.WriteAllText(this.filename, _code.GetText());
                this.Title = "RevitRubyShell " + this.filename;
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
                _code.SetText(File.ReadAllText(dlg.FileName));
                this.filename = dlg.FileName;
                this.Title = "RevitRubyShell " + this.filename;
            }      
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

        public static void SetText(this RichTextBox rtb, string value)
        {
            rtb.Document.Blocks.Clear();
            rtb.AppendText(value);
        }

        public static string GetText(this RichTextBox rtb)
        {
            if(!rtb.Selection.IsEmpty)
                return rtb.Selection.Text;
            TextRange textRange = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            return textRange.Text;

        }    
    }

    #endregion
}