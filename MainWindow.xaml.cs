using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace WindowsAppMvp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // Select the first program by default
            ProgramComboBox.SelectedIndex = 0;
        }

        // List to keep track of processes for clean shutdown
        private readonly List<Process> _processes = new List<Process>();

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            // Get the selected program path
            if (ProgramComboBox.SelectedItem is ComboBoxItem item && item.Content is string program)
            {
                OpenExternalApp(program);
            }
        }

        private void OpenExternalApp(string programPath)
        {
            try
            {
                // Create a new tab
                var tabItem = new TabItem();
                tabItem.Header = CreateTabHeader(programPath, tabItem);

                // Create WindowsFormsHost with panel for embedding
                var host = new WindowsFormsHost();
                var panel = new Panel();
                panel.BackColor = System.Drawing.Color.Black;
                panel.Dock = DockStyle.Fill;
                host.Child = panel;

                tabItem.Content = host;
                AppTabControl.Items.Add(tabItem);
                tabItem.IsSelected = true;

                // Start process and embed window
                var process = StartProcessAndEmbed(programPath, panel);
                if (process != null)
                {
                    // Store process for later cleanup
                    tabItem.Tag = process;
                    _processes.Add(process);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open {programPath}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private object CreateTabHeader(string title, TabItem tabItem)
        {
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
            var text = new TextBlock { Text = title, Margin = new Thickness(0, 0, 5, 0) };
            var closeButton = new Button
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
            // Close process associated with tab
            if (tabItem.Tag is Process proc)
            {
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                }
                catch
                {
                    // ignore errors
                }
                _processes.Remove(proc);
            }
            AppTabControl.Items.Remove(tabItem);
        }

        private Process StartProcessAndEmbed(string programPath, Panel hostPanel)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo(programPath)
                {
                    WorkingDirectory = System.IO.Path.GetDirectoryName(programPath),
                    UseShellExecute = true
                }
            };
            process.Start();
            // Wait for the process to create its main window
            process.WaitForInputIdle();

            IntPtr mainHandle = process.MainWindowHandle;
            if (mainHandle == IntPtr.Zero)
            {
                // Could not get handle
                return null;
            }

            // Re-parent the external window to the panel
            NativeMethods.SetParent(mainHandle, hostPanel.Handle);

            // Change style to child and remove title bar and resizable frame
            int style = NativeMethods.GetWindowLong(mainHandle, NativeMethods.GWL_STYLE);
            style = style | NativeMethods.WS_CHILD;
            style = style & ~NativeMethods.WS_CAPTION & ~NativeMethods.WS_THICKFRAME;
            NativeMethods.SetWindowLong(mainHandle, NativeMethods.GWL_STYLE, style);

            // Resize the embedded window to fill the panel
            ResizeEmbeddedWindow(mainHandle, hostPanel);

            // Hook size changed to update window size
            hostPanel.SizeChanged += (s, e) =>
            {
                ResizeEmbeddedWindow(mainHandle, hostPanel);
            };

            return process;
        }

        private void ResizeEmbeddedWindow(IntPtr windowHandle, Panel hostPanel)
        {
            NativeMethods.MoveWindow(windowHandle, 0, 0, hostPanel.ClientSize.Width, hostPanel.ClientSize.Height, true);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            // Kill all processes on exit
            foreach (var proc in _processes)
            {
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                }
                catch
                {
                    // ignore errors
                }
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        public const int GWL_STYLE = -16;
        public const int WS_CHILD = 0x40000000;
        public const int WS_VISIBLE = 0x10000000;
        public const int WS_CAPTION = 0x00C00000;
        public const int WS_THICKFRAME = 0x00040000;
    }
}
