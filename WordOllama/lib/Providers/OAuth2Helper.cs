using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Threading.Tasks;

namespace WordOllama.lib.Providers
{
    public static class OAuth2Helper
    {
        /// <summary>
        /// Perform OAuth2 authorization code flow via system browser + local HTTP listener.
        /// Returns the access token on success, or null on failure.
        /// </summary>
        public static async Task<string> AuthenticateAsync(string providerName, string clientId, string authUrl, string tokenUrl, string scope, string redirectPort = "51889")
        {
            string redirectUri = $"http://localhost:{redirectPort}/callback/";
            string state = Guid.NewGuid().ToString("N");

            // Build authorization URL
            string authorizeUrl = $"{authUrl}?" +
                $"client_id={Uri.EscapeDataString(clientId)}&" +
                $"response_type=code&" +
                $"redirect_uri={Uri.EscapeDataString(redirectUri)}&" +
                $"scope={Uri.EscapeDataString(scope)}&" +
                $"state={Uri.EscapeDataString(state)}&" +
                $"access_type=offline&" +
                $"prompt=consent";

            // Start local HTTP listener
            using (var listener = new HttpListener())
            {
                listener.Prefixes.Add($"http://localhost:{redirectPort}/");
                listener.Start();

                // Open browser
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = authorizeUrl,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    // Fallback: try to open via cmd
                    try
                    {
                        Process.Start("cmd", $"/c start {authorizeUrl}");
                    }
                    catch { }
                }

                // Wait for callback with timeout (120 seconds)
                var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(120));
                string authCode = null;

                try
                {
                    var contextTask = listener.GetContextAsync();
                    var completedTask = await Task.WhenAny(contextTask, Task.Delay(120_000, cts.Token));

                    if (completedTask == contextTask)
                    {
                        var context = contextTask.Result;
                        var request = context.Request;

                        // Verify state
                        string returnedState = request.QueryString["state"];
                        if (returnedState == state)
                        {
                            authCode = request.QueryString["code"];
                        }

                        // Return success HTML
                        string responseString = "<html><body><h3>认证成功! 请关闭此窗口。</h3></body></html>";
                        byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
                        context.Response.ContentType = "text/html; charset=utf-8";
                        context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                        context.Response.OutputStream.Close();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"OAuth2 callback error: {ex.Message}");
                    return null;
                }

                if (string.IsNullOrEmpty(authCode))
                    return null;

                // Exchange code for token
                return await ExchangeCodeForTokenAsync(tokenUrl, clientId, authCode, redirectUri, state);
            }
        }

        private static async Task<string> ExchangeCodeForTokenAsync(string tokenUrl, string clientId, string code, string redirectUri, string state)
        {
            try
            {
                using (var httpClient = new System.Net.Http.HttpClient())
                {
                    var content = new System.Net.Http.FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("code", code),
                        new KeyValuePair<string, string>("client_id", clientId),
                        new KeyValuePair<string, string>("redirect_uri", redirectUri),
                        new KeyValuePair<string, string>("grant_type", "authorization_code"),
                    });

                    var response = await httpClient.PostAsync(tokenUrl, content);
                    if (response.IsSuccessStatusCode)
                    {
                        string body = await response.Content.ReadAsStringAsync();
                        var json = Newtonsoft.Json.Linq.JObject.Parse(body);
                        return json["access_token"]?.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OAuth2 token exchange error: {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// Google OAuth2 for Gemini API. Returns an API key (in this case, the access token
        /// can be used as x-goog-api-key for some Google Cloud projects).
        /// </summary>
        public static async Task<string> GoogleOAuthAsync()
        {
            // Note: For Gemini, most users use API keys directly.
            // OAuth2 is needed for enterprise Google Cloud projects.
            // This requires a Google Cloud project with the "Generative Language API" enabled.
            return await AuthenticateAsync(
                providerName: "Google",
                clientId: "", // User needs to provide their own OAuth2 client ID
                authUrl: "https://accounts.google.com/o/oauth2/v2/auth",
                tokenUrl: "https://oauth2.googleapis.com/token",
                scope: "https://www.googleapis.com/auth/generative-language.retriever"
            );
        }
    }
}
