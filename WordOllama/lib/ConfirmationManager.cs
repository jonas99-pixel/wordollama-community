using System;
using System.Collections.Generic;
using WordOllama.Forms;

namespace WordOllama.lib
{
    /// <summary>
    /// 确认管理器 — 为 Agent 的各种确认弹窗提供会话级和永久级免打扰。
    /// </summary>
    public static class ConfirmationManager
    {
        // 确认点标识
        public const string KeyPlanConfirmation = "plan_confirmation";
        public const string KeyDiffReview = "diff_review";
        public const string KeyErrorRecovery = "error_recovery";

        // 会话级免打扰（内存中，重启后重置）
        private static readonly HashSet<string> _sessionOptOuts = new HashSet<string>();
        private static readonly HashSet<string> _sessionExternalPermissions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>检查指定确认点是否需要弹窗询问。</summary>
        public static bool ShouldPrompt(string key)
        {
            if (_sessionOptOuts.Contains(key)) return false;
            if (IsPermanentlyOptedOut(key)) return false;
            return true;
        }

        /// <summary>本次会话不再询问。</summary>
        public static void OptOutSession(string key) => _sessionOptOuts.Add(key);

        /// <summary>永久不再询问（保存到设置）。</summary>
        public static void OptOutPermanently(string key)
        {
            _sessionOptOuts.Add(key);
            SavePermanentOptOut(key);
        }

        /// <summary>清除所有会话级免打扰。</summary>
        public static void ClearSessionOptOuts()
        {
            _sessionOptOuts.Clear();
            _sessionExternalPermissions.Clear();
        }

        public static bool HasExternalPermissionForSession(string permissionKey) =>
            !string.IsNullOrWhiteSpace(permissionKey) && _sessionExternalPermissions.Contains(permissionKey);

        public static void GrantExternalPermissionForSession(string permissionKey)
        {
            if (!string.IsNullOrWhiteSpace(permissionKey))
                _sessionExternalPermissions.Add(permissionKey);
        }

        public static void ClearExternalPermissionsForSession() => _sessionExternalPermissions.Clear();

        private static bool IsPermanentlyOptedOut(string key)
        {
            return key switch
            {
                KeyPlanConfirmation => SharedSettings.SuppressPlanConfirmation,
                KeyDiffReview => SharedSettings.SuppressDiffReviewPrompt,
                KeyErrorRecovery => SharedSettings.SuppressErrorRecoveryPrompt,
                _ => false
            };
        }

        private static void SavePermanentOptOut(string key)
        {
            switch (key)
            {
                case KeyPlanConfirmation: SharedSettings.SuppressPlanConfirmation = true; break;
                case KeyDiffReview: SharedSettings.SuppressDiffReviewPrompt = true; break;
                case KeyErrorRecovery: SharedSettings.SuppressErrorRecoveryPrompt = true; break;
            }
        }
    }
}
