using Newtonsoft.Json.Linq;
using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    /// <summary>
    /// 读取表格内容 — 返回 Markdown 格式的表格数据。
    /// </summary>
    public class ReadTableTool : IAgentTool
    {
        public string Name => "read_table";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => false;
        public bool IsConcurrencySafe => true;
        public int MaxResultSizeChars => 10000;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["table_index"] = new JObject { ["type"] = "integer" }
            },
            ["required"] = new JArray("table_index")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int index = parameters["table_index"]?.Value<int>() ?? 1;
            if (index < 1 || index > doc.Tables.Count)
                return AgentToolText.F("I18n_AgentResultTableIndexOutOfRange", index, doc.Tables.Count);

            try
            {
                Word.Table table = doc.Tables[index];
                var sb = new StringBuilder();
                sb.AppendLine(AgentToolText.F("I18n_AgentResultTableHeader", index, table.Rows.Count, table.Columns.Count));
                sb.AppendLine();

                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    var row = table.Rows[r];
                    var cells = new System.Collections.Generic.List<string>();
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        string cellText = row.Cells[c].Range.Text?.Replace("\r", "").Replace("\a", "").Trim() ?? "";
                        cells.Add(cellText);
                    }
                    sb.AppendLine("| " + string.Join(" | ", cells) + " |");
                    if (r == 1) // 分隔符行
                    {
                        var sep = new System.Collections.Generic.List<string>();
                        for (int c = 1; c <= table.Columns.Count; c++)
                            sep.Add("---");
                        sb.AppendLine("| " + string.Join(" | ", sep) + " |");
                    }
                }

                DocumentAgent.FocusOnRange(table.Range, app);
                return sb.ToString().Trim();
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultReadTableFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 插入表格 — 在文档中创建新表格。
    /// </summary>
    public class InsertTableTool : IAgentTool
    {
        public string Name => "insert_table";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["rows"] = new JObject { ["type"] = "integer" },
                ["columns"] = new JObject { ["type"] = "integer" },
                ["header_row"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("rows", "columns")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int rows = parameters["rows"]?.Value<int>() ?? 2;
            int columns = parameters["columns"]?.Value<int>() ?? 2;
            string headerJson = parameters["header_row"]?.ToString() ?? "";

            if (rows < 1 || columns < 1)
                return AgentToolText.T("I18n_AgentResultTableDimensionsRequired");

            try
            {
                // 在文档末尾插入
                Word.Range rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                Word.Table table = doc.Tables.Add(rng, rows, columns);
                table.Borders.Enable = 1;

                // 设置表头
                if (!string.IsNullOrEmpty(headerJson))
                {
                    try
                    {
                        var headers = JArray.Parse(headerJson);
                        for (int c = 0; c < headers.Count && c < columns; c++)
                        {
                            table.Cell(1, c + 1).Range.Text = headers[c].ToString();
                        }
                        table.Rows[1].Range.Font.Bold = 1;
                        // 表头底纹
                        table.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                    }
                    catch
                    {
                        // JSON 解析失败则忽略
                    }
                }

                DocumentAgent.FocusOnRange(table.Range, app);
                return AgentToolText.F("I18n_AgentResultTableInserted", rows, columns);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultInsertTableFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 表格插入行 — 在指定表格中插入新行。
    /// </summary>
    public class TableInsertRowTool : IAgentTool
    {
        public string Name => "table_insert_row";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["table_index"] = new JObject { ["type"] = "integer" },
                ["after_row"] = new JObject { ["type"] = "integer" }
            },
            ["required"] = new JArray("table_index")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int tableIdx = parameters["table_index"]?.Value<int>() ?? 1;
            if (tableIdx < 1 || tableIdx > doc.Tables.Count)
                return AgentToolText.F("I18n_AgentResultTableIndexOutOfRangeSimple", tableIdx);

            try
            {
                Word.Table table = doc.Tables[tableIdx];
                object missing = System.Type.Missing;

                if (parameters["after_row"] != null)
                {
                    int afterRow = parameters["after_row"].Value<int>();
                    if (afterRow >= 1 && afterRow <= table.Rows.Count)
                    {
                        table.Rows[afterRow].Select();
                        app.Selection.InsertRowsBelow(1);
                    }
                    else
                    {
                        table.Rows.Add(ref missing);
                    }
                }
                else
                {
                    table.Rows.Add(ref missing);
                }

                return AgentToolText.F("I18n_AgentResultTableRowInserted", tableIdx, table.Rows.Count);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultInsertRowFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 设置表格单元格内容。
    /// </summary>
    public class TableSetCellTool : IAgentTool
    {
        public string Name => "table_set_cell";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["table_index"] = new JObject { ["type"] = "integer" },
                ["row"] = new JObject { ["type"] = "integer" },
                ["column"] = new JObject { ["type"] = "integer" },
                ["text"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("table_index", "row", "column", "text")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            int tableIdx = parameters["table_index"]?.Value<int>() ?? 1;
            int row = parameters["row"]?.Value<int>() ?? 1;
            int col = parameters["column"]?.Value<int>() ?? 1;
            string text = parameters["text"]?.ToString() ?? "";

            if (tableIdx < 1 || tableIdx > doc.Tables.Count)
                return AgentToolText.F("I18n_AgentResultTableIndexOutOfRangeSimple", tableIdx);

            try
            {
                Word.Table table = doc.Tables[tableIdx];
                if (row < 1 || row > table.Rows.Count || col < 1 || col > table.Columns.Count)
                    return AgentToolText.F("I18n_AgentResultCellOutOfRange", row, col, table.Rows.Count, table.Columns.Count);

                Word.Cell cell = table.Cell(row, col);
                cell.Range.Text = text;
                // 去掉单元格末尾的段落标记
                cell.Range.Paragraphs[1].Range.Font.Size = cell.Range.Font.Size;

                DocumentAgent.FocusOnRange(cell.Range, app);
                return AgentToolText.F("I18n_AgentResultTableCellSet", tableIdx, row, col);
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultSetCellFailed", ex.Message);
            }
        }
    }

    /// <summary>
    /// 更新/插入目录。
    /// </summary>
    public class UpdateTocTool : IAgentTool
    {
        public string Name => "update_toc";
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => true;
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 0;
        public JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject
            {
                ["action"] = new JObject { ["type"] = "string" }
            },
            ["required"] = new JArray("action")
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            string action = parameters["action"]?.ToString()?.ToLower() ?? "update";

            try
            {
                if (action == "update")
                {
                    // 更新所有 TOC 域
                    int count = 0;
                    foreach (Word.Field field in doc.Fields)
                    {
                        if (field.Type == Word.WdFieldType.wdFieldTOC)
                        {
                            field.Update();
                            count++;
                        }
                    }
                    return count > 0
                        ? AgentToolText.F("I18n_AgentResultTocUpdated", count)
                        : AgentToolText.T("I18n_AgentResultTocNotFound");
                }
                else if (action == "insert")
                {
                    // 在文档开头插入目录
                    Word.Range rng = doc.Range(0, 0);
                    rng.InsertParagraphBefore();
                    rng = doc.Range(0, 0);

                    // 插入"目录"标题
                    rng.Text = AgentToolText.T("I18n_AgentResultTocHeading");
                    rng.Font.Size = 16;
                    rng.Font.Bold = 1;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    rng.InsertParagraphAfter();

                    // 插入 TOC 域
                    Word.Range tocRange = doc.Range(doc.Paragraphs[2].Range.Start - 1, doc.Paragraphs[2].Range.Start - 1);
                    doc.Fields.Add(tocRange, Word.WdFieldType.wdFieldEmpty,
                        @"TOC \o ""1-3"" \h \z \u", false);

                    DocumentAgent.FocusOnRange(doc.Range(0, 0), app);
                    return AgentToolText.T("I18n_AgentResultTocInserted");
                }
                else
                {
                    return AgentToolText.F("I18n_AgentResultUnsupportedTocAction", action);
                }
            }
            catch (Exception ex)
            {
                return AgentToolText.F("I18n_AgentResultTocFailed", ex.Message);
            }
        }
    }
}
