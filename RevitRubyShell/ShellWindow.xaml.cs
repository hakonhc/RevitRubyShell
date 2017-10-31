using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Autodesk.Revit.UI;
using Microsoft.Scripting.Hosting;
using TextBox = System.Windows.Controls.TextBox;

namespace RevitRubyShell
{
    public partial class ShellWindow : IDockablePaneProvider
    {

        private ScriptEngine rubyEngine;
        private IronRuby.Runtime.RubyContext rubyContext;
        private string filename;
        private RevitRubyShellApplication myapp;

        public TextBox History => history;

        public TextBox Output => output;

        public TextBox Code => code;

        public GridSplitter ConsoleSplitter => consoleSplitter;

        public TextBoxBuffer OutputBuffer { get; internal set; }
                
        public ShellWindow()
        {
            InitializeComponent();
            myapp = RevitRubyShellApplication.RevitRubyShell;

            this.Loaded += (s, e) =>
                {                    
                    var defaultScripts = RevitRubyShellApplication.GetSettings().Root.Descendants("DefaultScript").ToArray();
                    var lastCode = myapp.LastCode;
                    if (string.IsNullOrEmpty(lastCode))
                    {
                        Code.Text = defaultScripts.Any() ? defaultScripts.First().Value.Replace("\n", "\r\n") : "";
                    }
                    else
                    {
                        Code.Text = lastCode;
                    }

                    OutputBuffer = new TextBoxBuffer(output);

                    // Initialize IronRuby
                    rubyEngine = myapp.RubyEngine;
                    rubyContext = (IronRuby.Runtime.RubyContext)Microsoft.Scripting.Hosting.Providers.HostingHelpers.GetLanguageContext(rubyEngine);          
                    

                    // redirect stdout to the output window
                    rubyContext.StandardOutput = OutputBuffer;                
                };

            this.Unloaded += (s, e) =>
                {
                    myapp.LastCode = Code.Text;
                };
        }


        /// <summary>
        /// Runs all code from a TextBox if there is no selection, otherwise
        /// just runs the selection.
        /// </summary>
        public void RunCode()
        {
            var code = Code.Text;
            RevitRubyShellApplication.RevitRubyShell.Queue.Enqueue(application =>
                {
                    myapp.RubyScope.SetVariable("_app", application);
                    string output;
                    var result = myapp.ExecuteCode(code, out output);
                    if (result)
                    {
                        OutputBuffer.Write(output);
                        // add the code to the history
                        history.AppendText(string.Format("{0}\n# {1}", code, output));
                    }
                    else
                    {
                        OutputBuffer.Write(output);
                    }
                });
        }

        //Save code buffer
        private void save_code(object sender, RoutedEventArgs e)
        {
            // Configure save file dialog box
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = filename ?? "command.rb"; // Default file name
            dlg.DefaultExt = ".rb"; // Default file extension
            dlg.Filter = "Ruby code (.rb)|*.rb"; // Filter files by extension

            // Process save file dialog box results
            if (dlg.ShowDialog() == true)
            {
                // Save document
                this.filename = dlg.FileName;
                File.WriteAllText(this.filename, Code.Text);
                this.Title = "RevitRubyShell " + this.filename;
            }
        }

        //Open rb file
        private void open_code(object sender, RoutedEventArgs e)
        {

            // Configure open file dialog box
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = filename ?? "command.rb"; // Default file name
            dlg.DefaultExt = ".rb"; // Default file extension
            dlg.Filter = "Ruby code (.rb)|*.rb"; // Filter files by extension

            // Process open file dialog box results
            if (dlg.ShowDialog() == true)
            {
                // Open document
                code.Text = File.ReadAllText(dlg.FileName);
                this.filename = dlg.FileName;
                this.Title = "RevitRubyShell " + this.filename;
            }      
        }

        private void run_code(object sender, RoutedEventArgs e)
        {
            RunCode();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
                {
                    DockPosition = DockPosition.Bottom
                };
        }

        private void _code_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.IsCtrl(Key.Enter))
            {
                RunCode();
                e.Handled = true;
            }
            else if (e.IsCtrl(Key.W))
            {
                Output.Clear();
                e.Handled = true;
            }
            else if (e.IsCtrl(Key.S))
            {
                save_code(null, null);
                e.Handled = true;
            }
        }
    }

    /// <summary>
    /// Simple TextBox Buffer class
    /// </summary>
    public class TextBoxBuffer
    {
        private readonly TextBox box;

        public TextBoxBuffer(TextBox t)
        {
            box = t;
        }

        public void Write(string str)
        {
            box.Dispatcher.BeginInvoke((Action)(() =>
            {
                box.AppendText(str);
                box.ScrollToEnd();
            }));
        }
    }

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
}