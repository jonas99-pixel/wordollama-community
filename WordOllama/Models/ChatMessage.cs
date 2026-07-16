using System;

namespace WordOllama.Models
{
    /// <summary>
    /// 聊天消息角色枚举。
    /// </summary>
    public enum ChatRole
    {
        System,
        User,
        Assistant
    }

    /// <summary>
    /// 聊天消息模型。
    /// </summary>
    public class ChatMessage : ObservableModel
    {
        private ChatRole _role = ChatRole.User;
        public ChatRole Role { get => _role; set => SetProperty(ref _role, value); }

        private string _content = "";
        public string Content { get => _content; set => SetProperty(ref _content, value); }

        private DateTime _timestamp = DateTime.Now;
        public DateTime Timestamp { get => _timestamp; set => SetProperty(ref _timestamp, value); }

        /// <summary>
        /// 转换为 Ollama API 所需的字典格式。
        /// </summary>
        public System.Collections.Generic.Dictionary<string, string> ToApiFormat()
        {
            return new System.Collections.Generic.Dictionary<string, string>
            {
                { "role", Role.ToString().ToLower() },
                { "content", Content }
            };
        }
    }
}
