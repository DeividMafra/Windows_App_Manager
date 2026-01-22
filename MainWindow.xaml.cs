using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Threading;
using System.Windows.Forms;
using System.Windows.Input;

namespace WindowsAppMvp
{
    public record ProgramEntry(string Title, string Command);

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadPrograms();
        }

        private readonly List<Process> _processes = new List<Process>();
        private const string ProgramsFileName = "programs.json";
        private List<ProgramEntry> _programs = new();

        private void LoadPrograms()
        {
            try
            {
                string path = @"C:\Dev\App_Manager\programs.json";
                if (System.IO.File.Exists(path))
                {
                    var json = System.IO.File.ReadAllText(path);
                    _programs = System.Text.Json.JsonSerializer.Deserialize<List<ProgramEntry>>(json) ?? new List<ProgramEntry>();
                }
                else
                {
                    _programs = new List<ProgramEntry>
                    {
                        new ProgramEntry("Notepad", "notepad.exe"),
                        new ProgramEntry("Calculator", "calc.exe"),
                        new ProgramEntry("Paint", "mspaint.exe")
                    };
                }

                MenuListBox.ItemsSource = _programs;
                if (_programs.Count > 0) MenuListBox.SelectedIndex = 0;
            }
            catch
            {
                _programs = new List<ProgramEntry>
                {
                    new ProgramEntry("Notepad", "notepad.exe"),
                    new ProgramEntry("Calculator", "calc.exe"),
                    new ProgramEntry("Paint", "mspaint.exe")
                };
                MenuListBox.ItemsSource = _programs;
                MenuListBox.SelectedIndex = 0;
            }
        }

        private void SavePrograms()
        {
            try
            {
                string path = @"C:\Dev\App_Manager\programs.json";
                var json = System.Text.Json.JsonSerializer.Serialize(_programs, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(path, json);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Failed saving programs list: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void MenuListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MenuListBox.SelectedItem is ProgramEntry entry && !string.IsNullOrWhiteSpace(entry.Command))
            {
                await OpenExternalAppAsync(entry.Command, entry.Title);
            }
        }

        private static (string file, string args) ParseProgramAndArgs(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return (input, string.Empty);
            input = input.Trim();
            if (input.StartsWith("\""))
            {
                int endQuote = input.IndexOf('"', 1);
                if (endQuote > 0)
                {
                    string file = input.Substring(1, endQuote - 1);
                    string args = input.Substring(endQuote + 1).Trim();
                    return (file, args);
                }
            }

            int firstSpace = input.IndexOf(' ');
            if (firstSpace < 0) return (input, string.Empty);
            string fileUnquoted = input.Substring(0, firstSpace);
            string remainingArgs = input.Substring(firstSpace + 1).Trim();
            return (fileUnquoted, remainingArgs);
        }

        private async Task OpenExternalAppAsync(string programCommand, string title)
        {
            TabItem tabItem = null;
            WindowsFormsHost host = null;
            System.Windows.Forms.Panel panel = null;

            try
            {
                Dispatcher.Invoke(() =>
                {
                    tabItem = new TabItem();
                    tabItem.Header = CreateTabHeader(title, tabItem);

                    host = new WindowsFormsHost();
                    panel = new System.Windows.Forms.Panel
                    {
                        BackColor = System.Drawing.Color.Black,
                        Dock = DockStyle.Fill
                    };
                    host.Child = panel;

                    tabItem.Content = host;
                    AppTabControl.Items.Add(tabItem);
                    tabItem.IsSelected = true;
                });

                var (file, args) = ParseProgramAndArgs(programCommand);

                var process = await StartProcessAndEmbedAsync(file, args, panel).ConfigureAwait(false);
                if (process != null)
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (tabItem != null)
                        {
                            tabItem.Tag = process;
                        }
                    });
                    _processes.Add(process);
                }
                else
                {
                    if (tabItem != null)
                    {
                        Dispatcher.Invoke(() => CloseTab(tabItem));
                    }
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show($"Failed to open {title}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private async Task<Process?> StartProcessAndEmbedAsync(string exePath, string arguments, System.Windows.Forms.Panel hostPanel)
        {
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments,
                WorkingDirectory = System.IO.Path.GetDirectoryName(exePath) ?? string.Empty,
                UseShellExecute = false
            };

            var process = new Process { StartInfo = psi };
            process.Start();

            var sw = Stopwatch.StartNew();
            const int timeoutMs = 5000;
            while (process.MainWindowHandle == IntPtr.Zero && !process.HasExited && sw.ElapsedMilliseconds < timeoutMs)
            {
                await Task.Delay(100).ConfigureAwait(false);
            }

            if (process.HasExited || process.MainWindowHandle == IntPtr.Zero)
            {
                try { if (!process.HasExited) process.Kill(); } catch { }
                return null;
            }

            IntPtr mainHandle = process.MainWindowHandle;

            Dispatcher.Invoke(() =>
            {
                NativeMethods.SetParent(mainHandle, hostPanel.Handle);

                long style = NativeMethods.GetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE).ToInt64();
                style &= ~(NativeMethods.WS_CAPTION | NativeMethods.WS_THICKFRAME);
                style |= (NativeMethods.WS_CHILD | NativeMethods.WS_VISIBLE);
                NativeMethods.SetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE, new IntPtr(style));

                NativeMethods.SetWindowPos(mainHandle, IntPtr.Zero, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height,
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_FRAMECHANGED);

                hostPanel.SizeChanged += (s, e) =>
                {
                    NativeMethods.SetWindowPos(mainHandle, IntPtr.Zero, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height,
                        NativeMethods.SWP_NOZORDER | NativeMethods.SWP_SHOWWINDOW);
                };
            });

            return process;
        }

        private object CreateTabHeader(string title, TabItem tabItem)
        {
            var headerPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
            var text = new TextBlock { Text = title, Margin = new Thickness(0, 0, 5, 0) };
            var closeButton = new System.Windows.Controls.Button
            {
                Content = "×",
                Width = 16,
                Height = 16,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
            };
            closeButton.Click += (s, e) => CloseTab(tabItem);

            headerPanel.Children.Add(text);
            headerPanel.Children.Add(closeButton);
            return headerPanel;
        }

        private void CloseTab(TabItem tabItem)
        {
            if (tabItem == null) return;
            if (tabItem.Tag is Process proc)
            {
                try
                {
                    if (!proc.HasExited) proc.Kill();
                }
                catch { }
                _processes.Remove(proc);
            }
            AppTabControl.Items.Remove(tabItem);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            foreach (var proc in _processes)
            {
                try
                {
                    if (!proc.HasExited) proc.Kill();
                }
                catch { }
            }
        }
    }
}