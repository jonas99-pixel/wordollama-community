using System;
using System.Security.Cryptography;
using System.Text;

namespace WordOllama.lib
{
    /// <summary>
    /// Protects locally persisted secrets for the current Windows user.
    /// Stored values are intentionally not portable between Windows accounts.
    /// </summary>
    internal static class SecretProtector
    {
        private const string Prefix = "dpapi:";
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("WordOllama.LocalSecret.v1");

        public static string Protect(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            byte[] plain = Encoding.UTF8.GetBytes(value);
            byte[] encrypted = ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser);
            return Prefix + Convert.ToBase64String(encrypted);
        }

        public static string UnprotectOrLegacy(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            if (!value.StartsWith(Prefix, StringComparison.Ordinal)) return value;

            try
            {
                byte[] encrypted = Convert.FromBase64String(value.Substring(Prefix.Length));
                byte[] plain = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch (CryptographicException)
            {
                return string.Empty;
            }
            catch (FormatException)
            {
                return string.Empty;
            }
        }

        public static bool IsProtected(string value)
        {
            return !string.IsNullOrEmpty(value) && value.StartsWith(Prefix, StringComparison.Ordinal);
        }
    }
}
