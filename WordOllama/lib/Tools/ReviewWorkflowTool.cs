using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib.Tools
{
    public sealed class ReviewWorkflowTool : IAgentTool, ICancellableAgentTool
    {
        private readonly Func<string, JObject, Word.Document, CancellationToken, Task<string>> _workflow;
        public string Name { get; }
        public ReviewWorkflowTool(string name, Func<string, JObject, Word.Document, CancellationToken, Task<string>> workflow)
        { Name = name; _workflow = workflow ?? throw new ArgumentNullException(nameof(workflow)); }
        public string Description => AgentToolText.Description(Name);
        public bool IsWriteOperation => Name == "apply_review_revisions";
        public bool IsConcurrencySafe => false;
        public int MaxResultSizeChars => 12000;
        public JObject ParameterSchema
        {
            get
            {
                var properties = new JObject();
                var required = new JArray();
                if (Name == "review_document")
                {
                    properties["kind"] = new JObject { ["type"] = "string", ["enum"] = new JArray("suggestions", "issues"), ["description"] = "Default suggestions. Use suggestions when tracked proofreading edits are requested; issues produces findings." };
                    properties["instruction"] = new JObject { ["type"] = "string", ["description"] = "Proofreading instructions for the complete document." };
                    properties["create_copy"] = new JObject { ["type"] = "boolean", ["description"] = "Defaults true. Preserve the original file by reviewing a separate native copy." };
                    properties["copy_path"] = new JObject { ["type"] = "string", ["description"] = "Optional new absolute path, with the original format. Existing files cannot be overwritten." };
                }
                if (Name == "resume_review")
                {
                    properties["session_id"] = new JObject { ["type"] = "string", ["description"] = "Saved session ID returned by list_review_sessions. Reopen its document/copy before resuming." };
                    required.Add("session_id");
                }
                if (Name == "save_review_document")
                    properties["export_report"] = new JObject { ["type"] = "boolean", ["description"] = "Defaults true. Save complete Markdown and JSON reports beside the reviewed document." };
                return new JObject { ["type"] = "object", ["properties"] = properties, ["required"] = required, ["additionalProperties"] = false };
            }
        }
        public string Execute(JObject parameters, Word.Document doc, Word.Application app) => throw new InvalidOperationException("Use the cancellable review workflow.");
        public Task<string> ExecuteAsync(JObject parameters, Word.Document doc, Word.Application app, CancellationToken cancellationToken)
            => _workflow(Name, parameters ?? new JObject(), doc, cancellationToken);
    }
}
