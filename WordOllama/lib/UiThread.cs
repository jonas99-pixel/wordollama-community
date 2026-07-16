using System;
using System.Windows.Threading;

namespace WordOllama.lib
{
    /// <summary>
    /// UI 线程安全工具。
    /// 统一处理跨线程 Dispatcher 访问，避免 InvalidOperationException。
    /// </summary>
    public static class UiThread
    {
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
