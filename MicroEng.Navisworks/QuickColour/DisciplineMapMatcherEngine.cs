using System;
using System.Text.RegularExpressions;

namespace MicroEng.Navisworks.QuickColour
{
    internal static class DisciplineMapMatcherEngine
    {
        public static bool IsMatch(string input, string type, string pattern)
        {
            input = input ?? "";
            pattern = pattern ?? "";

            var t = (type ?? "exact").Trim().ToLowerInvariant();
            switch (t)
            {
                case "exact":
                    return string.Equals(input, pattern, StringComparison.OrdinalIgnoreCase);

                case "contains":
                    return input.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0;

                case "startswith":
                case "starts_with":
                case "start":
                    return input.StartsWith(pattern, StringComparison.OrdinalIgnoreCase);

                case "endswith":
                case "ends_with":
                case "end":
                    return input.EndsWith(pattern, StringComparison.OrdinalIgnoreCase);

                case "wildcard":
                    return WildcardMatch(input, pattern);

                case "regex":
                    return Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                default:
                    return string.Equals(input, pattern, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool WildcardMatch(string input, string wildcard)
        {
            var rx = "^" + Regex.Escape(wildcard ?? "")
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";
            return Regex.IsMatch(input ?? "", rx, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
