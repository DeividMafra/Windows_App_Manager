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
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadPrograms();
        }

        private readonly List<Process> _processes = new List<Process>();
        private const string ProgramsFileName = "programs.json";
        private List<string> _programs = new();

        private void LoadPrograms()
        {
            try
            {
                // Always read from the fixed development path
                string path = @"C:\Dev\App_Manager\programs.json";
                if (System.IO.File.Exists(path))
                {
                    var json = System.IO.File.ReadAllText(path);
                    _programs = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
                }
                else
                {
                    _programs = new List<string> { "notepad.exe", "calc.exe", "mspaint.exe" };
                }

                MenuListBox.ItemsSource = _programs;
                if (_programs.Count > 0) MenuListBox.SelectedIndex = 0;
            }
            catch
            {
                // fall back to defaults
                _programs = new List<string> { "notepad.exe", "calc.exe", "mspaint.exe" };
                MenuListBox.ItemsSource = _programs;
                MenuListBox.SelectedIndex = 0;
            }
        }

        private void SavePrograms()
        {
            try
            {
                // keep persistence consistent with LoadPrograms path
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
            if (MenuListBox.SelectedItem is string program && !string.IsNullOrWhiteSpace(program))
            {
                await OpenExternalAppAsync(program);
            }
        }

        private static (string file, string args) ParseProgramAndArgs(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return (input, string.Empty);
            input = input.Trim();
            if (input.StartsWith("\""))
            {
                // quoted executable path
                int endQuote = input.IndexOf('"', 1);
                if (endQuote > 0)
                {
                    string file = input.Substring(1, endQuote - 1);
                    string args = input.Substring(endQuote + 1).Trim();
                    return (file, args);
                }
            }

            // unquoted: split on first space
            int firstSpace = input.IndexOf(' ');
            if (firstSpace < 0) return (input, string.Empty);
            string fileUnquoted = input.Substring(0, firstSpace);
            string remainingArgs = input.Substring(firstSpace + 1).Trim();
            return (fileUnquoted, remainingArgs);
        }

        private async Task OpenExternalAppAsync(string programPath)
        {
            TabItem tabItem = null;
            WindowsFormsHost host = null;
            System.Windows.Forms.Panel panel = null;

            try
            {
                // Create UI objects on the UI thread
                Dispatcher.Invoke(() =>
                {
                    tabItem = new TabItem();
                    tabItem.Header = CreateTabHeader(programPath, tabItem);

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

                var (file, args) = ParseProgramAndArgs(programPath);

                var process = await StartProcessAndEmbedAsync(file, args, panel).ConfigureAwait(false);
                if (process != null)
                {
                    // Tag and track process on UI thread
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
                    System.Windows.MessageBox.Show($"Failed to open {programPath}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

            // Wait for main window handle with timeout (don't block UI thread)
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

            // Set parent and change styles on UI thread (hostPanel.Handle must be accessed from the thread that owns it)
            Dispatcher.Invoke(() =>
            {
                NativeMethods.SetParent(mainHandle, hostPanel.Handle);

                // Update styles: remove caption/frame, set child & visible
                long style = NativeMethods.GetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE).ToInt64();
                style &= ~(NativeMethods.WS_CAPTION | NativeMethods.WS_THICKFRAME);
                style |= (NativeMethods.WS_CHILD | NativeMethods.WS_VISIBLE);
                NativeMethods.SetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE, new IntPtr(style));

                // Force window to update style and show
                NativeMethods.SetWindowPos(mainHandle, IntPtr.Zero, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height,
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_FRAMECHANGED);

                // Handle resizing
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