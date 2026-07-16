namespace WordOllama.lib
{
    public static class Loc
    {
        public static string T(string key)
        {
            string value = WordOllama.Properties.Resources.ResourceManager.GetString(key);
            return string.IsNullOrEmpty(value) ? key : value;
        }

        public static string F(string key, params object[] args)
        {
            return string.Format(T(key), args);
        }
    }
}
