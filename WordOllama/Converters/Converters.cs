using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using WordOllama.lib;

namespace WordOllama.Converters
{
    /// <summary>
    /// 布尔值取反转换器。
    /// </summary>
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is bool b ? !b : value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is bool b ? !b : value;
        }
    }

    /// <summary>
    /// AgentTaskStatus 到显示字符串的转换器。
    /// </summary>
    public class TaskStatusConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Models.AgentTaskStatus status)
            {
                switch (status)
                {
                    case Models.AgentTaskStatus.Pending: return Loc.T("I18n_TaskStatusPending");
                    case Models.AgentTaskStatus.Running: return Loc.T("I18n_TaskStatusRunning");
                    case Models.AgentTaskStatus.Completed: return Loc.T("I18n_TaskStatusCompleted");
                    case Models.AgentTaskStatus.Failed: return Loc.T("I18n_TaskStatusFailed");
                    case Models.AgentTaskStatus.Cancelled: return Loc.T("I18n_TaskStatusCancelled");
                }
            }
            return value?.ToString() ?? "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    /// <summary>
    /// 进度百分比到显示格式的转换器。
    /// </summary>
    public class ProgressConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double progress)
                return $"{progress:F0}%";
            return "0%";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    /// <summary>
    /// 对象非 null 时返回 Visible，否则 Collapsed。
    /// 用于 ImageSource 绑定。
    /// </summary>
    public class NullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
