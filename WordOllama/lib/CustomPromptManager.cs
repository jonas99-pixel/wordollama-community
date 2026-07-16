using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordOllama.lib
{
    public class CustomPrompt
    {
        public string Name { get; set; }
        public string Prompt { get; set; }
        public string Function { get; set; }
    }

    public static class CustomPromptManager
    {
        private static List<CustomPrompt> _prompts;

        public static List<CustomPrompt> LoadPrompts()
        {

                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordOllama", "OneClickCustomPrompts.json");
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    _prompts = JsonConvert.DeserializeObject<List<CustomPrompt>>(json);
                }
                else
                {
                    _prompts = new List<CustomPrompt>();
                }
            
            return _prompts;
        }

        public static CustomPrompt GetPromptByName(string name)
        {
            return LoadPrompts().FirstOrDefault(p => p.Name == name);
        }
    }
}
