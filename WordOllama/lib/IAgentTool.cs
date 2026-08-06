using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using WordOllama.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    /// <summary>
    /// Agent 工具接口 — 所有可被 AI 调用的文档操作工具必须实现此接口。
    /// 通过 ToolRegistry 自动发现并注册，实现插件化扩展。
    /// </summary>
    public interface IAgentTool
    {
        /// <summary>工具唯一名称（对应 AI 输出中的 action 字段）。</summary>
        string Name { get; }

        /// <summary>工具描述，用于生成给大模型的 Prompt。</summary>
        string Description { get; }

        /// <summary>参数 Schema（JSON Schema 格式），用于原生 Tool Calling。</summary>
        JObject ParameterSchema { get; }

        /// <summary>是否为写操作（会自动开启修订模式）。</summary>
        bool IsWriteOperation { get; }

        /// <summary>是否并发安全。只读工具返回 true，可与其它安全工具并行执行。</summary>
        bool IsConcurrencySafe { get; }

        /// <summary>
        /// 工具结果的最大字符数。超过此值的结果将被截断。
        /// 设为 0 表示使用全局默认值 (2000)。设为 -1 表示不限制。
        /// </summary>
        int MaxResultSizeChars { get; }

        /// <summary>执行工具操作。所有 COM 操作通过传入的 doc 对象完成。</summary>
        string Execute(JObject parameters, Word.Document doc, Word.Application app);
    }

    public interface ICancellableAgentTool
    {
        Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken);
    }

    /// <summary>
    /// 工具注册表 — 启动时自动发现程序集内所有 IAgentTool 实现并注册。
    /// </summary>
    public class ToolRegistry
    {
        private readonly Dictionary<string, IAgentTool> _tools = new Dictionary<string, IAgentTool>();

        /// <summary>
        /// 自动发现并注册所有 IAgentTool 实现。
        /// </summary>
        public void DiscoverAndRegister(bool includeExternalTools = false)
        {
            var toolType = typeof(IAgentTool);
            var asm = typeof(IAgentTool).Assembly;

            foreach (var type in asm.GetTypes())
            {
                if (type.IsAbstract || type.IsInterface) continue;
                if (!toolType.IsAssignableFrom(type)) continue;

                try
                {
                    // 处理 DocumentAgentTool<T> 泛型基类
                    IAgentTool tool;
                    if (type.BaseType != null && type.BaseType.IsGenericType &&
                        type.BaseType.GetGenericTypeDefinition() == typeof(DocumentAgentTool<>))
                    {
                        var innerType = type.BaseType.GetGenericArguments()[0];
                        var genericBase = typeof(DocumentAgentTool<>).MakeGenericType(innerType);
                        tool = (IAgentTool)System.Activator.CreateInstance(genericBase);
                    }
                    else
                    {
                        tool = (IAgentTool)System.Activator.CreateInstance(type);
                    }

                    if (!includeExternalTools && ToolExecutionPolicy.RequiresExternalToolPermission(tool))
                        continue;

                    _tools[tool.Name] = tool;
                }
                catch { }
            }
        }

        /// <summary>手动注册工具（用于测试或特殊场景）。</summary>
        public void Register(IAgentTool tool)
        {
            _tools[tool.Name] = tool;
        }

        public IAgentTool GetTool(string name)
        {
            _tools.TryGetValue(name, out var tool);
            return tool;
        }

        public IEnumerable<IAgentTool> GetAllTools() => _tools.Values;

        /// <summary>
        /// 生成工具描述文本（用于 ReAct Prompt）。
        /// </summary>
        public string BuildToolDescriptions()
        {
            var sb = new System.Text.StringBuilder();
            int i = 1;
            foreach (var tool in _tools.Values)
            {
                string effects = ToolExecutionPolicy.DescribeEffects(
                    ToolExecutionPolicy.GetEffects(tool),
                    AgentToolText.UseEnglish);
                sb.AppendLine($"{i}. {tool.Description} [Effects: {effects}]");
                i++;
            }
            return sb.ToString();
        }

        /// <summary>
        /// 生成 OpenAI / Ollama 兼容的 tools JSON Schema（用于原生 Tool Calling）。
        /// </summary>
        public JArray BuildToolSchemas()
        {
            var schemas = new JArray();
            foreach (var tool in _tools.Values)
            {
                schemas.Add(new JObject
                {
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["parameters"] = AgentToolText.LocalizeSchema(tool.Name, tool.ParameterSchema) ?? new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject(),
                            ["required"] = new JArray()
                        }
                    }
                });
            }
            return schemas;
        }
    }

    /// <summary>
    /// 泛型工具基类 — 简化常见工具的创建。
    /// 子类只需实现 ExecuteLogic，基类处理接口契约。
    /// </summary>
    public abstract class DocumentAgentTool<T> : IAgentTool where T : class
    {
        public abstract string Name { get; }
        public abstract string Description { get; }
        public virtual bool IsWriteOperation => false;
        public virtual bool IsConcurrencySafe => !IsWriteOperation;
        public virtual int MaxResultSizeChars => 0; // 0 = use global default
        public virtual JObject ParameterSchema => new JObject
        {
            ["type"] = "object",
            ["properties"] = new JObject(),
            ["required"] = new JArray()
        };

        public string Execute(JObject parameters, Word.Document doc, Word.Application app)
        {
            return ExecuteLogic(parameters, doc, app);
        }

        protected abstract string ExecuteLogic(JObject parameters, Word.Document doc, Word.Application app);
    }
}
