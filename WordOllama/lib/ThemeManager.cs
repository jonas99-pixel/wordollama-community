using WordOllama.lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using System.Windows.Forms;
using System.Windows.Interop;
using MaterialDesignThemes.Wpf;

namespace WordOllama.lib
{
    public static class ThemeManager
    {
        public static event EventHandler ThemeChanged;

        public static void ApplyTheme(ResourceDictionary resources)
        {
            var theme = resources.GetTheme();
            if (AppConfig.Instance.ThemeType == "Dark")
            {
                theme.SetBaseTheme(BaseTheme.Dark);
            }
            else
            {
                theme.SetBaseTheme(BaseTheme.Light);
            }
            resources.SetTheme(theme);

            // 触发主题变化事件
            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }
    }

    public static class HybridWindowManager
    {
        private static Dictionary<Type, object> openWindows = new Dictionary<Type, object>();

        public static void ShowWindow<T>() where T : new()
        {
            if (openWindows.TryGetValue(typeof(T), out object existingWindow))
            {
                if (existingWindow is Form formWindow && !formWindow.IsDisposed)
                {
                    formWindow.Activate();
                    if (formWindow.WindowState == FormWindowState.Minimized)
                    {
                        formWindow.WindowState = FormWindowState.Normal;
                    }
                }
                else if (existingWindow is Window wpfWindow && !wpfWindow.IsDisposed())
                {
                    wpfWindow.Activate();
                    if (wpfWindow.WindowState == WindowState.Minimized)
                    {
                        wpfWindow.WindowState = WindowState.Normal;
                    }
                }
                else
                {
                    CreateAndShowNewWindow<T>();
                }
            }
            else
            {
                CreateAndShowNewWindow<T>();
            }
        }

        private static void CreateAndShowNewWindow<T>() where T : new()
        {
            T newWindow = new T();
            openWindows[typeof(T)] = newWindow;

            if (newWindow is Form formWindow)
            {
                formWindow.FormClosed += (sender, args) => openWindows.Remove(typeof(T));
                formWindow.Show();
            }
            else if (newWindow is Window wpfWindow)
            {
                wpfWindow.ShowInTaskbar = false;
                wpfWindow.Closed += (sender, args) => openWindows.Remove(typeof(T));
                // Set Word main window as owner to group taskbar icon and auto-close on Word exit
                try
                {
                    var wih = new WindowInteropHelper(wpfWindow);
                    wih.Owner = Process.GetCurrentProcess().MainWindowHandle;
                }
                catch { }
                wpfWindow.Show();
            }
        }

        /// <summary>
        /// 关闭所有已打开的窗口，在 Word 关闭时调用以防止进程残留。
        /// </summary>
        public static void CloseAllWindows()
        {
            var keysToRemove = new List<Type>();
            foreach (var kv in openWindows)
            {
                if (kv.Value is Window wpf && !wpf.IsDisposed())
                {
                    try { wpf.Close(); } catch { }
                    keysToRemove.Add(kv.Key);
                }
                else if (kv.Value is Form form && !form.IsDisposed)
                {
                    try { form.Close(); } catch { }
                    keysToRemove.Add(kv.Key);
                }
            }
            foreach (var key in keysToRemove)
                openWindows.Remove(key);
        }
    }

    public static class WindowExtensions
    {
        public static bool IsDisposed(this Window window)
        {
            return PresentationSource.FromVisual(window) == null;
        }
    }
}
