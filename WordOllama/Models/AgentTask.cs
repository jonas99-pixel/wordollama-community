using System;
using System.Collections.ObjectModel;

namespace WordOllama.Models
{
    /// <summary>
    /// Agent 任务状态枚举。
    /// </summary>
    public enum AgentTaskStatus
    {
        Pending,
        Running,
        Completed,
        Failed,
        Cancelled
    }

    /// <summary>
    /// 单个 Agent 子步骤。
    /// </summary>
    public class AgentStep : ObservableModel
    {
        private int _stepNumber;
        public int StepNumber { get => _stepNumber; set => SetProperty(ref _stepNumber, value); }

        private string _description = "";
        public string Description { get => _description; set => SetProperty(ref _description, value); }

        private string _prompt = "";
        public string Prompt { get => _prompt; set => SetProperty(ref _prompt, value); }

        private string _result = "";
        public string Result { get => _result; set => SetProperty(ref _result, value); }

        private AgentTaskStatus _status = AgentTaskStatus.Pending;
        public AgentTaskStatus Status { get => _status; set => SetProperty(ref _status, value); }

        private string _errorMessage = "";
        public string ErrorMessage { get => _errorMessage; set => SetProperty(ref _errorMessage, value); }

        public DateTime? StartedAt { get; set; }
        public DateTime? CompletedAt { get; set; }
    }

    /// <summary>
    /// Agent 多步骤任务。
    /// </summary>
    public class AgentTask : ObservableModel
    {
        private string _taskId = Guid.NewGuid().ToString("N").Substring(0, 8);
        public string TaskId { get => _taskId; set => SetProperty(ref _taskId, value); }

        private string _userRequirement = "";
        public string UserRequirement { get => _userRequirement; set => SetProperty(ref _userRequirement, value); }

        private AgentTaskStatus _status = AgentTaskStatus.Pending;
        public AgentTaskStatus Status { get => _status; set => SetProperty(ref _status, value); }

        private ObservableCollection<AgentStep> _steps = new ObservableCollection<AgentStep>();
        public ObservableCollection<AgentStep> Steps { get => _steps; set => SetProperty(ref _steps, value); }

        private int _currentStepIndex = -1;
        public int CurrentStepIndex { get => _currentStepIndex; set => SetProperty(ref _currentStepIndex, value); }

        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime? CompletedAt { get; set; }

        /// <summary>
        /// 获取当前正在执行的步骤。
        /// </summary>
        public AgentStep CurrentStep =>
            CurrentStepIndex >= 0 && CurrentStepIndex < Steps.Count ? Steps[CurrentStepIndex] : null;

        /// <summary>
        /// 总步骤数。
        /// </summary>
        public int TotalSteps => Steps.Count;

        /// <summary>
        /// 已完成步骤数。
        /// </summary>
        public int CompletedSteps
        {
            get
            {
                int count = 0;
                foreach (var s in Steps)
                    if (s.Status == AgentTaskStatus.Completed) count++;
                return count;
            }
        }
    }
}
