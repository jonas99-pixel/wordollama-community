using System;

namespace WordOllama.lib
{
    public enum ToolCallingMode
    {
        Auto,
        ReAct,
        Native
    }

    public static class ProviderCompatibility
    {
        public static ToolCallingMode ParseToolCallingMode(string value)
        {
            if (string.Equals(value, "ReAct", StringComparison.OrdinalIgnoreCase)) return ToolCallingMode.ReAct;
            if (string.Equals(value, "Native", StringComparison.OrdinalIgnoreCase)) return ToolCallingMode.Native;
            return ToolCallingMode.Auto;
        }

        public static string ToStorageValue(ToolCallingMode mode) => mode.ToString();
    }
}
