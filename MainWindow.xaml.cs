using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Windows.Threading;

namespace WindowsAppMvp
{
    /// <summary>
    /// Represents a program entry in the launcher. This record holds a display title,
    /// the command to run and an optional working directory.
    /// </summary>
    public record ProgramEntry(string Title, string Command, string? StartIn = null);

    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// This window acts as a simple launcher/window manager. Programs defined in
    /// a JSON file are displayed in a ListBox. Double‑clicking on an item launches
    /// the corresponding executable and embeds its main window into a tab.
    /// </summary>
    public partial class MainWindow : Window
    {
        // Maintain a collection of running processes so we can clean them up on exit.
        private readonly List<Process> _processes = new();

        // Use an observable collection so changes propagate to the UI.
        private ObservableCollection<ProgramEntry> _programs = new();

        // Name of the JSON file that stores program definitions.
        private const string ProgramsFileName = "programs.json";

        // Compute the path to the JSON file in the user's local app data folder.
        private static readonly string ProgramsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "App_Manager",
            ProgramsFileName);

        public MainWindow()
        {
            InitializeComponent();
            LoadPrograms();
        }

        /// <summary>
        /// Loads program definitions from disk. If no file is found a default set
        /// is created. Any deserialization errors fall back to defaults.
        /// </summary>
        private void LoadPrograms()
        {
            try
            {
                if (File.Exists(ProgramsFilePath))
                {
                    var json = File.ReadAllText(ProgramsFilePath);
                    var list = System.Text.Json.JsonSerializer.Deserialize<List<ProgramEntry>>(json);
                    _programs = list != null
                        ? new ObservableCollection<ProgramEntry>(list)
                        : GetDefaultPrograms();
                }
                else
                {
                    _programs = GetDefaultPrograms();
                }
            }
            catch
            {
                // If anything goes wrong during load, use defaults.
                _programs = GetDefaultPrograms();
            }

            // Bind the list to the UI.
            MenuListBox.ItemsSource = _programs;
            if (_programs.Count > 0)
            {
                MenuListBox.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Writes the current program definitions to disk. Ensures the directory exists
        /// before writing the file. Displays an error message if an exception occurs.
        /// </summary>
        private void SavePrograms()
        {
            try
            {
                var directory = Path.GetDirectoryName(ProgramsFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = System.Text.Json.JsonSerializer.Serialize(_programs, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(ProgramsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed saving programs list: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Builds a default set of program entries used when no JSON exists or loading fails.
        /// </summary>
        private static ObservableCollection<ProgramEntry> GetDefaultPrograms() =>
            new(new[]
            {
                new ProgramEntry("Notepad", "notepad.exe"),
                new ProgramEntry("Calculator", "calc.exe"),
                new ProgramEntry("Paint", "mspaint.exe")
            });

        /// <summary>
        /// Handles double‑click events in the program list. Launches the selected entry.
        /// </summary>
        private async void MenuListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MenuListBox.SelectedItem is ProgramEntry entry && !string.IsNullOrWhiteSpace(entry.Command))
            {
                await OpenExternalAppAsync(entry.Command, entry.Title, entry.StartIn);
            }
        }

        /// <summary>
        /// Parses a command string into an executable and a list of arguments.
        /// Handles quoted tokens so paths with spaces are treated as a single argument.
        /// </summary>
        private static (string file, List<string> args) ParseProgramAndArgs(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return (string.Empty, new List<string>());
            }

            input = input.Trim();
            var tokens = new List<string>();

            int i = 0;
            int n = input.Length;

            while (i < n)
            {
                // Skip whitespace.
                while (i < n && char.IsWhiteSpace(input[i]))
                    i++;
                if (i >= n)
                    break;

                string token;
                if (input[i] == '"')
                {
                    // Quoted token. Consume until the matching quote.
                    i++;
                    int start = i;
                    while (i < n && input[i] != '"')
                        i++;
                    token = input.Substring(start, i - start);
                    if (i < n && input[i] == '"')
                        i++;
                }
                else
                {
                    // Unquoted token. Consume until whitespace.
                    int start = i;
                    while (i < n && !char.IsWhiteSpace(input[i]))
                        i++;
                    token = input.Substring(start, i - start);
                }

                tokens.Add(token);
            }

            if (tokens.Count == 0)
            {
                return (string.Empty, new List<string>());
            }

            var file = tokens[0];
            var args = tokens.Skip(1).ToList();
            return (file, args);
        }

        /// <summary>
        /// Launches an external application, embeds its main window into a new tab and
        /// registers its lifetime so it can be closed automatically when the process exits.
        /// </summary>
        private async Task OpenExternalAppAsync(string programCommand, string title, string? startIn)
        {
            TabItem? tabItem = null;
            WindowsFormsHost? host = null;
            Panel? panel = null;

            try
            {
                // Create UI elements on the dispatcher thread first so bindings work.
                await Dispatcher.InvokeAsync(() =>
                {
                    tabItem = new TabItem();

                    host = new WindowsFormsHost();
                    panel = new Panel
                    {
                        BackColor = System.Drawing.Color.Black,
                        Dock = DockStyle.Fill
                    };
                    host.Child = panel;

                    tabItem.Content = host;
                    AppTabControl.Items.Add(tabItem);
                    tabItem.IsSelected = true;
                });

                // Assign header after the tab is in the visual tree so FindAncestor works.
                await Dispatcher.InvokeAsync(() =>
                {
                    if (tabItem != null)
                    {
                        tabItem.Header = CreateTabHeader(title, tabItem);
                    }
                });

                // Parse the command into file and argument list.
                var (file, args) = ParseProgramAndArgs(programCommand);

                // Start the process and embed it.
                var process = await StartProcessAndEmbedAsync(file, args, panel!, startIn).ConfigureAwait(false);

                if (process != null)
                {
                    // Store the process on the tab so we can terminate it when closing the tab.
                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (tabItem != null)
                        {
                            tabItem.Tag = process;
                        }
                    });

                    _processes.Add(process);

                    // When the process exits, close its tab automatically.
                    process.EnableRaisingEvents = true;
                    process.Exited += (_, __) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            CloseTab(tabItem!);
                        });
                    };
                }
                else
                {
                    // If the process could not start, remove the tab.
                    if (tabItem != null)
                    {
                        await Dispatcher.InvokeAsync(() => CloseTab(tabItem));
                    }
                }
            }
            catch (Exception ex)
            {
                // Display an error and clean up the tab if something goes wrong.
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"Failed to open {title}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    if (tabItem != null)
                    {
                        CloseTab(tabItem);
                    }
                });
            }
        }

        /// <summary>
        /// Starts a process given an executable and argument list, waits for its window
        /// to become available and embeds it into the supplied WinForms panel. If the
        /// process fails to start or display a window within the timeout, null is returned.
        /// </summary>
        private async Task<Process?> StartProcessAndEmbedAsync(string exePath, List<string> arguments, Panel hostPanel, string? workingDirectory = null)
        {
            if (string.IsNullOrWhiteSpace(exePath))
            {
                return null;
            }

            // Build the start info. Use ArgumentList rather than Arguments to avoid quoting issues.
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = false,
                WorkingDirectory = !string.IsNullOrWhiteSpace(workingDirectory)
                    ? workingDirectory
                    : (Path.GetDirectoryName(exePath) ?? string.Empty)
            };

            foreach (var arg in arguments)
            {
                psi.ArgumentList.Add(arg);
            }

            var process = new Process { StartInfo = psi };

            if (!process.Start())
            {
                return null;
            }

            // Optionally wait for the process to be ready for input.
            try
            {
                process.WaitForInputIdle(5000);
            }
            catch
            {
                // Some processes may not support WaitForInputIdle; ignore any exceptions.
            }

            // Wait for a window handle within a timeout.
            var sw = Stopwatch.StartNew();
            const int timeoutMs = 5000;
            while (process.MainWindowHandle == IntPtr.Zero && !process.HasExited && sw.ElapsedMilliseconds < timeoutMs)
            {
                await Task.Delay(100).ConfigureAwait(false);
            }

            if (process.HasExited || process.MainWindowHandle == IntPtr.Zero)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                    }
                }
                catch
                {
                    // Ignored: the process may already be gone.
                }
                return null;
            }

            // Embed the window into the panel on the UI thread.
            IntPtr mainHandle = process.MainWindowHandle;
            await Dispatcher.InvokeAsync(() =>
            {
                NativeMethods.SetParent(mainHandle, hostPanel.Handle);

                long style = NativeMethods.GetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE).ToInt64();
                style &= ~(NativeMethods.WS_CAPTION | NativeMethods.WS_THICKFRAME);
                style |= (NativeMethods.WS_CHILD | NativeMethods.WS_VISIBLE);
                NativeMethods.SetWindowLongPtr(mainHandle, NativeMethods.GWL_STYLE, new IntPtr(style));

                NativeMethods.SetWindowPos(mainHandle, IntPtr.Zero, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height,
                    NativeMethods.SWP_NOZORDER | NativeMethods.SWP_SHOWWINDOW | NativeMethods.SWP_FRAMECHANGED);

                hostPanel.SizeChanged += (_, __) =>
                {
                    NativeMethods.SetWindowPos(mainHandle, IntPtr.Zero, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height,
                        NativeMethods.SWP_NOZORDER | NativeMethods.SWP_SHOWWINDOW);
                };
            });

            return process;
        }

        /// <summary>
        /// Creates a header for each tab containing a title and a close button.
        /// </summary>
        private object CreateTabHeader(string title, TabItem tabItem)
        {
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            var text = new TextBlock { Text = title, Margin = new Thickness(0, 0, 5, 0) };
            var closeButton = new System.Windows.Controls.Button
            {
                Content = "×",
                Width = 16,
                Height = 16,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
            };
            closeButton.Click += (_, __) => CloseTab(tabItem);

            headerPanel.Children.Add(text);
            headerPanel.Children.Add(closeButton);
            return headerPanel;
        }

        /// <summary>
        /// Closes a tab and terminates the associated process if it is still running.
        /// </summary>
        private void CloseTab(TabItem tabItem)
        {
            if (tabItem == null)
            {
                return;
            }
            if (tabItem.Tag is Process proc)
            {
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                    proc.Dispose();
                }
                catch
                {
                    // ignored
                }
                _processes.Remove(proc);
            }
            AppTabControl.Items.Remove(tabItem);
        }

        /// <summary>
        /// When the window closes, ensure all started processes are terminated.
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            foreach (var proc in _processes.ToList())
            {
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                    proc.Dispose();
                }
                catch
                {
                    // ignored
                }
            }
            _processes.Clear();
        }
    }
}