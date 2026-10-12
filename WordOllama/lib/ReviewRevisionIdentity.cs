using System;
using System.Globalization;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOllama.lib
{
    public static class ReviewRevisionIdentity
    {
        public static string Build(string name, string mode, DateTime date, string session)
        {
            string result = string.IsNullOrWhiteSpace(name) ? "AI审批" : name.Trim();
            if (result.Length > 128 || result.Any(char.IsControl)) throw new ArgumentException(Loc.T("I18n_ReviewReviewerInvalid"));
            if (string.Equals(mode, "Date", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode, "Session", StringComparison.OrdinalIgnoreCase))
                result += "-" + date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (string.Equals(mode, "Session", StringComparison.OrdinalIgnoreCase))
                result += "-" + (string.IsNullOrEmpty(session) ? Guid.NewGuid().ToString("N") : session).Substring(0, Math.Min(8, (string.IsNullOrEmpty(session) ? 32 : session.Length)));
            return result;
        }

        public static IDisposable Enter(Word.Application application, string identity)
        {
            if (application == null || string.IsNullOrWhiteSpace(identity)) return new IdentityScope(null, null);
            return new IdentityScope(application, identity);
        }

        private sealed class IdentityScope : IDisposable
        {
            private Word.Application _application;
            private readonly string _name, _initials;
            private Word.Options _options;
            private bool? _useLocal;
            public IdentityScope(Word.Application application, string identity)
            {
                _application = application;
                if (application == null) return;
                _name = application.UserName; _initials = application.UserInitials;
                try
                {
                    _options = application.Options;
                    try { _useLocal = _options.UseLocalUserInfo; _options.UseLocalUserInfo = true; }
                    catch (System.Runtime.InteropServices.COMException) { }
                    application.UserName = identity; application.UserInitials = "AI";
                }
                catch { Dispose(); throw; }
            }
            public void Dispose()
            {
                var application = _application; _application = null;
                if (application == null) return;
                // Never hold this scope across an await or a user interaction.
                try { application.UserName = _name; }
                finally
                {
                    try { application.UserInitials = _initials; }
                    finally
                    {
                        try { if (_useLocal.HasValue) _options.UseLocalUserInfo = _useLocal.Value; }
                        finally { ComHelper.ReleaseCom(ref _options); }
                    }
                }
            }
        }
    }
}
