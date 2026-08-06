using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WordOllama.lib
{
    public static class ComHelper
    {
        /// <summary>
        /// 安全释放 COM 对象。用于循环中创建的临时 COM 对象释放。
        /// VSTO 插件中不需要对所有 COM 对象调用此方法——运行时已管理大部分生命周期。
        /// 仅在以下场景使用：(1) 循环中反复创建的临时对象；(2) 静态/长生命期引用。
        /// </summary>
        public static void ReleaseCom<T>(ref T obj) where T : class
        {
            if (obj != null)
            {
                try
                {
                    if (Marshal.IsComObject(obj))
                        Marshal.ReleaseComObject(obj);
                }
                catch { /* 对象已被释放或不可用 */ }
                obj = null;
            }
        }
    }

    public static class Utils
    {
        public static async Task<bool> OptimizeTextBoxAsync(TextBox textBox, Button trigger = null)
        {
            if (textBox == null) return false;
            string original = textBox.Text;
            if (trigger != null) trigger.IsEnabled = false;
            try
            {
                string optimized = await PromptOptimizer(original);
                if (string.IsNullOrWhiteSpace(optimized)) return false;
                UiThread.RunSync(textBox, () =>
                {
                    textBox.Text = optimized;
                    textBox.CaretIndex = textBox.Text.Length;
                    textBox.Focus();
                });
                return true;
            }
            finally
            {
                if (trigger != null)
                    UiThread.RunSync(trigger, () => trigger.IsEnabled = true);
            }
        }

        public static async Task<string> PromptOptimizer(string userPrompt)
        {
            if (string.IsNullOrWhiteSpace(userPrompt))
            {
                MessageBox.Show(WordOllama.Properties.Resources.请输入提示词, WordOllama.Properties.Resources.提示, MessageBoxButton.OK, MessageBoxImage.Warning);
                return string.Empty;
            }

            try
            {
                string responseLanguage = PromptTemplateService.ResolveLanguage(
                    AppConfig.Instance?.LanguageMode,
                    userPrompt);
                string finalPrompt = PromptTemplateService.BuildPromptOptimizerPrompt(userPrompt, responseLanguage);
                var response = await Globals.ThisAddIn.GenerateModelSeletor(finalPrompt);
                if (response != null)
                {
                    return response;
                }
                else
                {
                    return WordOllama.Properties.Resources.错误AI没有返回结果请检查;
                }

            } 
            catch (Exception ex) {
                string errorMessage = $"{WordOllama.Properties.Resources.优化提示词出错}: {ex.Message}";
                MessageBox.Show(errorMessage, WordOllama.Properties.Resources.错误, MessageBoxButton.OK, MessageBoxImage.Error);
                return string.Empty;
            }
        }
    }
}
