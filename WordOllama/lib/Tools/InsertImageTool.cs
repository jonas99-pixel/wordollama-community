using System;
using System.IO;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// 图片插入工具 — 在 Word 文档中插入图片。
    /// 支持本地文件路径或 base64 编码数据两种方式。
    /// 自动被 ToolRegistry.DiscoverAndRegister() 发现注册。
    /// </summary>
    public class InsertImageTool : IAgentTool
    {
        public string Name => "insert_image";
        public string Description => AgentToolText.Description(Name);

        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => !IsWriteOperation;
        public int MaxResultSizeChars => 0;

        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["file_path"] = new JObject { ["type"] = "string" },
                ["base64_data"] = new JObject { ["type"] = "string" },
                ["paragraph_index"] = new JObject { ["type"] = "integer" },
                ["width"] = new JObject { ["type"] = "number" }
            },
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string filePath = parameters["file_path"]?.ToString()?.Trim() ?? "";
            string base64Data = parameters["base64_data"]?.ToString()?.Trim() ?? "";
            int? paragraphIndex = parameters["paragraph_index"]?.Value<int>();
            double? width = parameters["width"]?.Value<double>();

            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(base64Data))
                return AgentToolText.T("I18n_AgentResultImageSourceRequired");

            string tempFile = null;

            try
            {
                // 处理 base64 数据 → 临时文件
                if (!string.IsNullOrEmpty(base64Data))
                {
                    string cleanBase64 = base64Data;
                    int commaIndex = cleanBase64.IndexOf(',');
                    if (commaIndex >= 0)
                        cleanBase64 = cleanBase64.Substring(commaIndex + 1);

                    byte[] imageBytes = Convert.FromBase64String(cleanBase64);
                    tempFile = Path.GetTempFileName() + ".png";
                    File.WriteAllBytes(tempFile, imageBytes);
                    filePath = tempFile;
                }

                if (!File.Exists(filePath))
                    return AgentToolText.F("I18n_AgentResultImageFileNotFound", filePath);

                // 确定插入位置
                Word.Range insertRange;
                if (paragraphIndex.HasValue && paragraphIndex.Value >= 1 && paragraphIndex.Value <= doc.Paragraphs.Count)
                {
                    insertRange = doc.Paragraphs[paragraphIndex.Value].Range;
                }
                else
                {
                    insertRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                }

                // 插入图片
                Word.InlineShape shape = doc.InlineShapes.AddPicture(filePath, Type.Missing, Type.Missing, insertRange);

                if (width.HasValue && width.Value > 0)
                {
                    shape.Width = (float)width.Value;
                }

                DocumentAgent.FocusOnRange(insertRange, app);

                string location = paragraphIndex.HasValue
                    ? AgentToolText.F("I18n_AgentResultParagraphLocation", paragraphIndex.Value)
                    : AgentToolText.T("I18n_AgentResultDocumentEnd");
                string sizeInfo = width.HasValue ? AgentToolText.F("I18n_AgentResultImageWidth", width.Value) : AgentToolText.T("I18n_AgentResultOriginalSize");

                return AgentToolText.F("I18n_AgentResultImageInserted", location, sizeInfo);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultInsertImageFailed", ex.Message);
            }
            finally
            {
                if (tempFile != null)
                {
                    try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                }
            }
        }
    }
}
