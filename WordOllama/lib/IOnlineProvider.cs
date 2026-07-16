using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;

namespace WordOllama.lib
{
    public interface IOnlineProvider
    {
        string ProviderType { get; }
        Task<string> GenerateAsync(string prompt, string imageBase64 = null, double? temperature = null);
        Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools = null);
        Task StreamAsync(string prompt, Action<string> processChunk);
        Task<List<string>> FetchModelsAsync();
    }

    public interface ICancellableOnlineProvider
    {
        Task<string> ChatAsync(List<Dictionary<string, string>> chatHistory, JArray tools, CancellationToken cancellationToken);
    }
}
