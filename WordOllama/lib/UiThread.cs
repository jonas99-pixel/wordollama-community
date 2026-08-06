using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;

namespace WordOllama.lib
{
    /// <summary>
    /// UI 线程安全工具。
    /// 统一处理跨线程 Dispatcher 访问，避免 InvalidOperationException。
    /// </summary>
    public static class UiThread
    {
        private static readonly object OfficeSync = new object();
        private static Control _officeInvoker;
        private static int _officeThreadId;

        /// <summary>
        /// Captures Word's STA thread. Call once from ThisAddIn_Startup before any
        /// background Agent work is started.
        /// </summary>
        public static void InitializeOfficeThread()
        {
            lock (OfficeSync)
            {
                if (_officeInvoker != null && !_officeInvoker.IsDisposed) return;
                _officeThreadId = Thread.CurrentThread.ManagedThreadId;
                _officeInvoker = new Control();
                _officeInvoker.CreateControl();
                IntPtr handle = _officeInvoker.Handle;
            }
        }

        public static bool CheckOfficeAccess() =>
            _officeThreadId != 0 && Thread.CurrentThread.ManagedThreadId == _officeThreadId;

        public static void RunOfficeSync(Action action)
        {
            if (action == null) return;
            Control invoker = _officeInvoker;
            if (invoker == null || invoker.IsDisposed || CheckOfficeAccess())
            {
                action();
                return;
            }
            invoker.Invoke(action);
        }

        public static T RunOfficeSync<T>(Func<T> func)
        {
            if (func == null) return default(T);
            Control invoker = _officeInvoker;
            if (invoker == null || invoker.IsDisposed || CheckOfficeAccess())
                return func();
            return (T)invoker.Invoke(func);
        }

        public static Task RunOfficeAsync(Action action)
        {
            if (action == null) return Task.CompletedTask;
            Control invoker = _officeInvoker;
            if (invoker == null || invoker.IsDisposed || CheckOfficeAccess())
            {
                action();
                return Task.CompletedTask;
            }

            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            invoker.BeginInvoke(new Action(() =>
            {
                try { action(); completion.TrySetResult(true); }
                catch (Exception ex) { completion.TrySetException(ex); }
            }));
            return completion.Task;
        }

        public static void DisposeOfficeThread()
        {
            lock (OfficeSync)
            {
                Control invoker = _officeInvoker;
                _officeInvoker = null;
                _officeThreadId = 0;
                if (invoker == null) return;
                try
                {
                    if (invoker.InvokeRequired) invoker.BeginInvoke(new Action(invoker.Dispose));
                    else invoker.Dispose();
                }
                catch { }
            }
        }

        /// <summary>
        /// 安全执行：如果在 UI 线程则直接执行，否则调度到 UI 线程。
        /// </summary>
        public static void Run(DispatcherObject target, Action action)
        {
            if (target.Dispatcher.CheckAccess())
                action();
            else
                target.Dispatcher.BeginInvoke(action);
        }

        /// <summary>
        /// 安全执行：指定 Dispatcher。
        /// </summary>
        public static void Run(Dispatcher dispatcher, Action action)
        {
            if (dispatcher.CheckAccess())
                action();
            else
                dispatcher.BeginInvoke(action);
        }

        /// <summary>
        /// 安全执行并等待结果。
        /// </summary>
        public static T RunSync<T>(DispatcherObject target, Func<T> func)
        {
            if (target.Dispatcher.CheckAccess())
                return func();
            else
                return (T)target.Dispatcher.Invoke(func);
        }

        /// <summary>
        /// 安全执行并等待。
        /// </summary>
        public static void RunSync(DispatcherObject target, Action action)
        {
            if (target.Dispatcher.CheckAccess())
                action();
            else
                target.Dispatcher.Invoke(action);
        }
    }
}
