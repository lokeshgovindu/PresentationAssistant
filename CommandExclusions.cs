using System;
using System.Collections.Generic;
using System.Linq;

namespace PresentationAssistant
{
    /// <summary>
    /// The user's own list of commands not to announce, from the "Excluded Commands"
    /// setting. Patterns are matched case-insensitively against the canonical command
    /// name and may end in <c>*</c> to match a prefix, e.g. <c>Edit.Line*</c>.
    /// </summary>
    internal sealed class CommandExclusions
    {
        public static readonly CommandExclusions Empty = new CommandExclusions(null);

        /// <summary>Separators accepted when the patterns are written as a single string.</summary>
        private static readonly char[] Separators = { ';', ',', '\r', '\n' };

        private readonly HashSet<string> _exact = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _prefixes = new List<string>();

        public CommandExclusions(IEnumerable<string> patterns)
        {
            if (patterns == null) return;

            foreach (var raw in patterns)
            {
                var pattern = (raw ?? string.Empty).Trim();
                if (pattern.Length == 0) continue;

                if (pattern.EndsWith("*", StringComparison.Ordinal))
                {
                    var prefix = pattern.Substring(0, pattern.Length - 1);
                    // A bare "*" would hide every command, which is never what someone
                    // means by an exclusion list.
                    if (prefix.Length > 0) _prefixes.Add(prefix);
                }
                else
                {
                    _exact.Add(pattern);
                }
            }
        }

        public bool IsEmpty => _exact.Count == 0 && _prefixes.Count == 0;

        /// <summary>Splits a settings string such as "Edit.Line*; View.Output" into patterns.</summary>
        public static string[] Split(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? new string[0]
                : value.Split(Separators, StringSplitOptions.RemoveEmptyEntries)
                       .Select(p => p.Trim())
                       .Where(p => p.Length > 0)
                       .ToArray();
        }

        /// <summary>Joins patterns back into the single-line form shown in the options grid.</summary>
        public static string Join(IEnumerable<string> patterns)
        {
            return patterns == null ? string.Empty : string.Join("; ", patterns);
        }

        public bool IsExcluded(string actionId)
        {
            if (string.IsNullOrEmpty(actionId)) return false;
            if (_exact.Contains(actionId)) return true;

            for (var i = 0; i < _prefixes.Count; i++)
            {
                if (actionId.StartsWith(_prefixes[i], StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }
    }
}
