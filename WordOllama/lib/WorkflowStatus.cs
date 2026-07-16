using System;

namespace WordOllama.lib
{
    public enum WorkflowState
    {
        Idle,
        Starting,
        Thinking,
        Acting,
        Observing,
        Reviewing,
        WaitingForUser,
        Completed,
        Cancelled,
        Failed
    }

    public sealed class WorkflowStatusSnapshot
    {
        public WorkflowState State { get; }
        public string Detail { get; }
        public int Current { get; }
        public int Total { get; }
        public DateTime Timestamp { get; }

        public WorkflowStatusSnapshot(WorkflowState state, string detail, int current = 0, int total = 0)
        {
            State = state;
            Detail = detail ?? "";
            Current = Math.Max(0, current);
            Total = Math.Max(0, total);
            Timestamp = DateTime.Now;
        }
    }

    /// <summary>
    /// Shared status publisher for Agent, Review Workspace, and Silent Review.
    /// </summary>
    public sealed class WorkflowStatus
    {
        public WorkflowStatusSnapshot Current { get; private set; } =
            new WorkflowStatusSnapshot(WorkflowState.Idle, "");

        public event Action<WorkflowStatusSnapshot> Changed;

        public WorkflowStatusSnapshot Set(WorkflowState state, string detail, int current = 0, int total = 0)
        {
            Current = new WorkflowStatusSnapshot(state, detail, current, total);
            Changed?.Invoke(Current);
            return Current;
        }
    }
}
