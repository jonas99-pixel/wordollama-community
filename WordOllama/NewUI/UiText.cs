using System;
using WordOllama.lib;

namespace WordOllama.NewUI
{
    public static class UiText
    {
        public static string T(string key) => Loc.T(key);
        public static string F(string key, params object[] args) => Loc.F(key, args);

        public static string ReviewTab => Loc.T("I18n_UiReviewTab");
        public static string ChatTab => Loc.T("I18n_UiChatTab");
        public static string IssuesTab => Loc.T("I18n_UiIssuesTab");
        public static string AgentTitle => Loc.T("I18n_UiAgentTitle");
        public static string Stop => Loc.T("I18n_UiStop");
        public static string Thinking => Loc.T("I18n_UiThinking");
        public static string ConfirmPlan => Loc.T("I18n_UiConfirmPlan");
        public static string RevisePlan => Loc.T("I18n_UiRevisePlan");
        public static string NoIssues => Loc.T("I18n_UiNoIssues");
        public static string Clear => Loc.T("I18n_UiClear");
        public static string ReviewSelection => Loc.T("I18n_UiReviewSelection");
        public static string ReviewDocument => Loc.T("I18n_UiReviewDocument");
        public static string IssuesEmpty => Loc.T("I18n_UiIssuesEmpty");
        public static string Jump => Loc.T("I18n_UiJump");
        public static string Comment => Loc.T("I18n_UiComment");
        public static string Ignore => Loc.T("I18n_UiIgnore");
        public static string InputHint => Loc.T("I18n_UiInputHint");
        public static string RemoveImage => Loc.T("I18n_UiRemoveImage");
        public static string AgentEmptyTitle => Loc.T("I18n_UiAgentEmptyTitle");
        public static string AgentEmptyState => Loc.T("I18n_UiAgentEmptyState");
        public static string AgentExecutionDetails => Loc.T("I18n_UiAgentExecutionDetails");
        public static string AgentResultTitle => Loc.T("I18n_UiAgentResultTitle");

        public static string ReviewWorkspaceTitle => Loc.T("I18n_UiReviewWorkspaceTitle");
        public static string ReviewPaneTitle => Loc.T("I18n_UiReviewPaneTitle");
        public static string ReviewWorkspaceSubtitle => Loc.T("I18n_UiReviewWorkspaceSubtitle");
        public static string LoadSelection => Loc.T("I18n_UiLoadSelection");
        public static string LoadParagraphs => Loc.T("I18n_UiLoadParagraphs");
        public static string LoadDocument => Loc.T("I18n_UiLoadDocument");
        public static string ReviewScopeLabel => Loc.T("I18n_UiReviewScopeLabel");
        public static string ReviewScopeHint => Loc.T("I18n_UiReviewScopeHint");
        public static string ReviewScopeNone => Loc.T("I18n_UiReviewScopeNone");
        public static string ReviewGenerationLabel => Loc.T("I18n_UiReviewGenerationLabel");
        public static string ReviewBatchActions => Loc.T("I18n_UiReviewBatchActions");
        public static string GenerateAll => Loc.T("I18n_UiGenerateAll");
        public static string AcceptAll => Loc.T("I18n_UiAcceptAll");
        public static string InsertAll => Loc.T("I18n_UiInsertAll");
        public static string CommentAll => Loc.T("I18n_UiCommentAll");
        public static string SkipAll => Loc.T("I18n_UiSkipAll");
        public static string ProfileLabel => Loc.T("I18n_UiProfileLabel");
        public static string ProfileHint => Loc.T("I18n_UiProfileHint");
        public static string SaveProfile => Loc.T("I18n_UiSaveProfile");
        public static string EmptySuggestions => Loc.T("I18n_UiEmptySuggestions");
        public static string Generate => Loc.T("I18n_UiGenerate");
        public static string Focus => Loc.T("I18n_UiFocus");
        public static string AcceptReplace => Loc.T("I18n_UiAcceptReplace");
        public static string InsertBelow => Loc.T("I18n_UiInsertBelow");
        public static string Copy => Loc.T("I18n_UiCopy");
        public static string Skip => Loc.T("I18n_UiSkip");
        public static string Original => Loc.T("I18n_UiOriginal");
        public static string Suggestion => Loc.T("I18n_UiSuggestion");
        public static string CustomInstruction => Loc.T("I18n_UiCustomInstruction");
        public static string CustomInstructionHint => Loc.T("I18n_UiCustomInstructionHint");
        public static string DefaultReviewInstruction => Loc.T("I18n_UiDefaultReviewInstruction");
        public static string CancelReviewGeneration => Loc.T("I18n_UiCancelReviewGeneration");

        public static string ApiTab => Loc.T("I18n_UiApiTab");
        public static string ApiTitle => Loc.T("I18n_UiApiTitle");
        public static string ApiDescription => Loc.T("I18n_UiApiDescription");
        public static string RefreshApi => Loc.T("I18n_UiRefreshApi");
        public static string CopyApi => Loc.T("I18n_UiCopyApi");
        public static string ApiSafetyNote => Loc.T("I18n_UiApiSafetyNote");
        public static string SkillName => Loc.T("I18n_UiSkillName");
        public static string SkillDescription => Loc.T("I18n_UiSkillDescription");
        public static string SkillGoal => Loc.T("I18n_UiSkillGoal");
        public static string GenerateSkill => Loc.T("I18n_UiGenerateSkill");
        public static string SaveSkill => Loc.T("I18n_UiSaveSkill");
        public static string SkillDraft => Loc.T("I18n_UiSkillDraft");
        public static string ApiNameRequired => Loc.T("I18n_UiApiNameRequired");
        public static string ApiGoalRequired => Loc.T("I18n_UiApiGoalRequired");
        public static string ApiGenerating => Loc.T("I18n_UiApiGenerating");
        public static string ApiGenerationFailed => Loc.T("I18n_UiApiGenerationFailed");
        public static string ApiSkillSaved => Loc.T("I18n_UiApiSkillSaved");
        public static string ApiSkillExists => Loc.T("I18n_UiApiSkillExists");
        public static string ApiInvalidSkillName => Loc.T("I18n_UiApiInvalidSkillName");
        public static string ApiBuildFailed => Loc.T("I18n_UiApiBuildFailed");
        public static string ApiCopied => Loc.T("I18n_UiApiCopied");
        public static string ApiEmptyModel => Loc.T("I18n_UiApiEmptyModel");
        public static string ApiDraftRequired => Loc.T("I18n_UiApiDraftRequired");
    }
}
