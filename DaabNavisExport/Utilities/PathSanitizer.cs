using System;
using System.IO;
using System.Linq;
using System.Text;

namespace DaabNavisExport.Utilities
{
    internal static class PathSanitizer
    {
        private static readonly char[] InvalidFileNameChars = Path.GetInvalidFileNameChars();

        public static string ToSafeFileName(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "navisworks";
            }

            var text = value!;
            var builder = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                builder.Append(InvalidFileNameChars.Contains(ch) ? '_' : ch);
            }

            var sanitized = builder.ToString().Trim();
            return string.IsNullOrEmpty(sanitized) ? "navisworks" : sanitized;
        }

        public static string ToSafePath(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "";
            }

            var invalid = Path.GetInvalidPathChars();
            var sanitized = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            return sanitized;
        }
    }
}
