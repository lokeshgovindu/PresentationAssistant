using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// The set of themes the user can pick from: the built-ins below, plus anything in
    /// themes.json. A file entry whose name matches a built-in overrides it; any other
    /// name is a new theme.
    /// </summary>
    /// <remarks>
    /// Built-ins deliberately stay in code rather than being written out to themes.json,
    /// so improvements to them reach existing users instead of being shadowed by a stale
    /// copy on disk.
    /// </remarks>
    internal sealed class ThemeCatalog
    {
        /// <summary>Fallback when nothing else resolves. Also the baseline for brand new themes.</summary>
        public static readonly ThemePalette Fallback = new ThemePalette(
            ThemeNames.Classic,
            background:          Hex("#DCF2DC"),
            foreground:          Hex("#1E1E1E"),
            secondaryForeground: Hex("#5A6B5A"),
            border:              Hex("#7FA97F"),
            opacity:             0.85,
            isDark:              false);

        private static readonly IReadOnlyList<ThemePalette> BuiltIns = new ReadOnlyCollection<ThemePalette>(
            new List<ThemePalette>
            {
                // ---- light backgrounds ----
                Fallback,                                                   // Classic - the original overlay
                P("Blue",           "#DCEAF7", "#10243A", "#4B6C8C", "#2E6DA4", 0.85, false),
                P("Amber",          "#FDF1D6", "#3A2A05", "#8A6D24", "#C89A2B", 0.85, false),
                P("Teal",           "#D9F2EF", "#0C2B28", "#457F78", "#2A9D8F", 0.85, false),
                P("Rose",           "#FBE4EA", "#3A1020", "#8C5A6B", "#C9526E", 0.85, false),
                P("Lavender",       "#ECE3F7", "#241436", "#6C5088", "#7B4FA8", 0.85, false),
                P("Slate",          "#E8EAED", "#1F2328", "#5C636B", "#8A9299", 0.85, false),
                P("SolarizedLight", "#FDF6E3", "#073642", "#657B83", "#B58900", 0.90, false),

                // ---- dark backgrounds ----
                P("ClassicDark",    "#16241B", "#E4F2E7", "#93AE9B", "#4C9A63", 0.90, true),
                P("Purple",         "#2A1F3D", "#F0E9FA", "#AC97C9", "#7B4FA8", 0.90, true),
                P("Midnight",       "#0F1B2D", "#E6EDF7", "#8199B8", "#2E6DA4", 0.90, true),
                P("SolarizedDark",  "#002B36", "#EEE8D5", "#93A1A1", "#B58900", 0.90, true),
                P("Nord",           "#2E3440", "#ECEFF4", "#9AA5B5", "#5E81AC", 0.92, true),
                P("Dracula",        "#282A36", "#F8F8F2", "#A0A5C0", "#BD93F9", 0.92, true),
                P("Monokai",        "#272822", "#F8F8F2", "#A6A28C", "#A6E22E", 0.92, true),
                P("GruvboxDark",    "#282828", "#EBDBB2", "#A89984", "#D79921", 0.92, true),
                P("OneDark",        "#282C34", "#D7DAE0", "#9199A8", "#61AFEF", 0.92, true),
                P("TokyoNight",     "#1A1B26", "#C0CAF5", "#7E85A8", "#7AA2F7", 0.92, true),
                P("HighContrast",   "#000000", "#FFFFFF", "#FFFF00", "#FFFFFF", 1.00, true),
            });

        /// <summary>How often to re-stat themes.json, in milliseconds.</summary>
        private const int FileCheckIntervalMs = 1000;

        private static ThemeCatalog _current;
        private static DateTime _fileStamp;
        private static int _lastCheckTicks;
        private static bool _seedAttempted;

        private readonly Dictionary<string, ThemePalette> _byName;
        private readonly Dictionary<string, ThemeDefinition> _dynamicOverrides;
        private readonly ReadOnlyCollection<string> _names;

        private ThemeCatalog(
            IEnumerable<ThemePalette> palettes,
            Dictionary<string, ThemeDefinition> dynamicOverrides)
        {
            Palettes = new ReadOnlyCollection<ThemePalette>(palettes.ToList());
            _byName = Palettes.ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);
            _dynamicOverrides = dynamicOverrides;
            _names = new ReadOnlyCollection<string>(
                new[] { ThemeNames.Auto, ThemeNames.VisualStudio }
                    .Concat(Palettes.Select(p => p.Name))
                    .ToList());
        }

        /// <summary>Every theme, built-ins first, in presentation order.</summary>
        public IReadOnlyList<ThemePalette> Palettes { get; }

        /// <summary>
        /// Names for the options dropdown: the two dynamic entries first, then every
        /// palette.
        /// </summary>
        public IReadOnlyList<string> Names => _names;

        /// <summary>
        /// Bumped on every rebuild, so callers can cache a resolved palette and know when
        /// it went stale without comparing contents.
        /// </summary>
        public static int Version { get; private set; }

        /// <summary>
        /// The catalog, reloaded whenever themes.json has changed on disk so that editing
        /// the file shows up without restarting Visual Studio.
        /// </summary>
        /// <remarks>
        /// This is consulted once per announced command, so the file is only re-stat'ed at
        /// most once a second. Nobody edits a theme and gets back to the IDE faster than
        /// that, and it keeps file I/O off the per-command path.
        /// </remarks>
        public static ThemeCatalog Current
        {
            get
            {
                if (_current == null)
                {
                    Reload();
                }
                else if (unchecked(Environment.TickCount - _lastCheckTicks) >= FileCheckIntervalMs)
                {
                    _lastCheckTicks = Environment.TickCount;
                    if (ThemesFileStamp() != _fileStamp) Reload();
                }

                return _current;
            }
        }

        /// <summary>Drops the cache so the next access re-reads themes.json.</summary>
        public static void Invalidate() => _current = null;

        /// <summary>
        /// Writes <paramref name="definition"/> into themes.json, replacing any existing
        /// entry with the same name. This is how the options page persists colours edited
        /// in the UI.
        /// </summary>
        /// <remarks>
        /// Rewriting the file loses any comments in it. The documentation lives in the
        /// generated themes.reference.json, which is never rewritten, so nothing the user
        /// needs is lost - but a hand-annotated themes.json will come back tidied.
        /// </remarks>
        public static void SaveOverride(ThemeDefinition definition)
        {
            if (definition == null || string.IsNullOrWhiteSpace(definition.Name)) return;

            var entries = ReadDefinitions().ToList();
            entries.RemoveAll(d => SameName(d, definition.Name));
            entries.Add(definition);

            WriteDefinitions(entries);
        }

        /// <summary>Drops any themes.json entry for <paramref name="name"/>, reverting it to the built-in.</summary>
        public static void RemoveOverride(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;

            var entries = ReadDefinitions().ToList();
            if (entries.RemoveAll(d => SameName(d, name)) == 0) return;

            WriteDefinitions(entries);
        }

        /// <summary>Whether themes.json currently carries an entry for this theme.</summary>
        public static bool HasOverride(string name)
        {
            return !string.IsNullOrWhiteSpace(name) && ReadDefinitions().Any(d => SameName(d, name));
        }

        private static bool SameName(ThemeDefinition definition, string name)
        {
            return string.Equals(
                Canonical(definition?.Name),
                Canonical(name),
                StringComparison.OrdinalIgnoreCase);
        }

        private static void WriteDefinitions(IEnumerable<ThemeDefinition> definitions)
        {
            try
            {
                AppPaths.EnsureDataFolder();

                // Ignore nulls so an entry only lists the fields that were actually set.
                var json = JsonConvert.SerializeObject(definitions, Formatting.Indented,
                    new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                File.WriteAllText(AppPaths.ThemesFile, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{Product.Name}] Failed to write themes.json: {ex}");
            }
            finally
            {
                Invalidate();
            }
        }

        private static void Reload()
        {
            _fileStamp = ThemesFileStamp();
            _lastCheckTicks = Environment.TickCount;
            _current = Build();
            Version++;
        }

        public bool TryGet(string name, out ThemePalette palette)
        {
            palette = null;
            return !string.IsNullOrWhiteSpace(name) && _byName.TryGetValue(Canonical(name), out palette);
        }

        /// <summary>
        /// Maps the names earlier builds used onto the current ones, so a saved setting or
        /// a themes.json entry saying "Light"/"Dark" still lands on Classic/ClassicDark.
        /// </summary>
        private static string Canonical(string name)
        {
            var trimmed = (name ?? string.Empty).Trim();

            if (string.Equals(trimmed, ThemeNames.LegacyLight, StringComparison.OrdinalIgnoreCase))
                return ThemeNames.Classic;
            if (string.Equals(trimmed, ThemeNames.LegacyDark, StringComparison.OrdinalIgnoreCase))
                return ThemeNames.ClassicDark;

            return trimmed;
        }

        /// <summary>
        /// Applies a themes.json entry named after one of the computed themes (Auto,
        /// VisualStudio) on top of the palette that was just computed for it, so those
        /// two can be tuned like any other theme instead of silently ignoring the file.
        /// </summary>
        public ThemePalette ApplyDynamicOverride(string name, ThemePalette computed)
        {
            return _dynamicOverrides.TryGetValue(name, out var definition)
                ? definition.ToPalette(computed).Rename(computed.Name)
                : computed;
        }

        /// <summary>The palette Auto uses on a light shell.</summary>
        public ThemePalette LightDefault =>
            TryGet(ThemeNames.Classic, out var p) ? p : Palettes.FirstOrDefault(x => !x.IsDark) ?? Fallback;

        /// <summary>The palette Auto uses on a dark shell.</summary>
        public ThemePalette DarkDefault =>
            TryGet(ThemeNames.ClassicDark, out var p) ? p : Palettes.FirstOrDefault(x => x.IsDark) ?? Fallback;

        private static ThemeCatalog Build()
        {
            var ordered = BuiltIns.ToList();
            var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < ordered.Count; i++) index[ordered[i].Name] = i;

            var dynamicOverrides = new Dictionary<string, ThemeDefinition>(StringComparer.OrdinalIgnoreCase);

            foreach (var definition in ReadDefinitions())
            {
                if (string.IsNullOrWhiteSpace(definition.Name)) continue;

                var name = Canonical(definition.Name);

                // Auto and VisualStudio are computed rather than stored, so they are held
                // aside and merged onto the computed result instead of becoming palettes
                // of their own (which would double them up in the dropdown).
                if (string.Equals(name, ThemeNames.Auto, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, ThemeNames.VisualStudio, StringComparison.OrdinalIgnoreCase))
                {
                    dynamicOverrides[name] = definition;
                    continue;
                }

                if (index.TryGetValue(name, out var at))
                {
                    // Override a built-in in place, keeping its position in the list.
                    ordered[at] = definition.ToPalette(ordered[at]);
                }
                else
                {
                    ordered.Add(definition.ToPalette(Fallback.Rename(name)));
                    index[name] = ordered.Count - 1;
                }
            }

            return new ThemeCatalog(ordered, dynamicOverrides);
        }

        private static IEnumerable<ThemeDefinition> ReadDefinitions()
        {
            SeedThemesFileOnce();

            try
            {
                if (!File.Exists(AppPaths.ThemesFile)) return Enumerable.Empty<ThemeDefinition>();

                var json = File.ReadAllText(AppPaths.ThemesFile);
                if (string.IsNullOrWhiteSpace(json)) return Enumerable.Empty<ThemeDefinition>();

                // Comments are allowed - the seeded file uses them to document the format.
                return JsonConvert.DeserializeObject<List<ThemeDefinition>>(json)
                       ?? (IEnumerable<ThemeDefinition>)Enumerable.Empty<ThemeDefinition>();
            }
            catch (Exception ex)
            {
                // A malformed themes.json must leave the built-ins working.
                Debug.WriteLine($"[{Product.Name}] Failed to read themes.json: {ex}");
                return Enumerable.Empty<ThemeDefinition>();
            }
        }

        private static DateTime ThemesFileStamp()
        {
            try
            {
                return File.Exists(AppPaths.ThemesFile)
                    ? File.GetLastWriteTimeUtc(AppPaths.ThemesFile)
                    : DateTime.MinValue;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Makes sure themes.json exists and the generated reference listing is current.
        /// Safe to call repeatedly; used by the "Themes File" button on the options page.
        /// </summary>
        public static void EnsureAuthoringFiles()
        {
            _seedAttempted = true;
            WriteSeedIfMissing();
            WriteReferenceFile();
        }

        /// <summary>
        /// Creates themes.json the first time we look for it, so there is something to
        /// discover and copy from. Only attempted once per session, so deleting the file
        /// on purpose keeps it deleted.
        /// </summary>
        private static void SeedThemesFileOnce()
        {
            if (_seedAttempted) return;
            _seedAttempted = true;

            WriteSeedIfMissing();
            WriteReferenceFile();
        }

        private static void WriteSeedIfMissing()
        {
            try
            {
                if (File.Exists(AppPaths.ThemesFile)) return;

                AppPaths.EnsureDataFolder();
                File.WriteAllText(AppPaths.ThemesFile, SeedContents());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{Product.Name}] Failed to seed themes.json: {ex}");
            }
        }

        /// <summary>
        /// Writes every built-in theme out in full, in the exact shape themes.json uses,
        /// so editing a built-in is copy-paste-tweak rather than guessing hex codes.
        /// Deliberately generated from the pristine built-ins, not from the merged
        /// catalog, so it stays a stable starting point.
        /// </summary>
        private static void WriteReferenceFile()
        {
            try
            {
                AppPaths.EnsureDataFolder();

                var sb = new StringBuilder();
                sb.AppendLine("// PresentationAssistant - the built-in themes, in full.");
                sb.AppendLine("//");
                sb.AppendLine("// GENERATED, and rewritten every Visual Studio session: editing THIS file has no");
                sb.AppendLine("// effect. Copy an entry into themes.json and change it there - an entry whose");
                sb.AppendLine("// \"Name\" matches overrides the built-in, and you only need to keep the fields");
                sb.AppendLine("// you actually want to change.");
                sb.AppendLine("//");
                sb.AppendLine("// \"Auto\" and \"VisualStudio\" are computed from the running IDE rather than listed");
                sb.AppendLine("// here, but they can be tuned the same way.");
                sb.AppendLine("[");

                for (var i = 0; i < BuiltIns.Count; i++)
                {
                    var p = BuiltIns[i];
                    sb.Append("  { ");
                    sb.Append($"\"Name\": \"{p.Name}\", ");
                    sb.Append($"\"Background\": \"{Hex(p.BackgroundColor)}\", ");
                    sb.Append($"\"Foreground\": \"{Hex(p.ForegroundColor)}\", ");
                    sb.Append($"\"SecondaryForeground\": \"{Hex(p.SecondaryForegroundColor)}\", ");
                    sb.Append($"\"Border\": \"{Hex(p.BorderColor)}\", ");
                    sb.Append($"\"Opacity\": {p.Opacity.ToString("0.00", CultureInfo.InvariantCulture)}, ");
                    sb.Append($"\"IsDark\": {(p.IsDark ? "true" : "false")}");
                    sb.Append(" }");
                    if (i < BuiltIns.Count - 1) sb.Append(",");
                    sb.AppendLine();
                }

                sb.AppendLine("]");

                File.WriteAllText(AppPaths.ThemesReferenceFile, sb.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{Product.Name}] Failed to write themes.reference.json: {ex}");
            }
        }

        private static string Hex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        private static string SeedContents()
        {
            var builtIns = string.Join(", ", BuiltIns.Select(p => p.Name));

            return
$@"// PresentationAssistant custom themes.
//
// Each entry adds a theme, or overrides a built-in one when the name matches.
// Only ""Name"" is required - anything you leave out is inherited, and omitting
// ""SecondaryForeground"" derives it from the other two colours.
//
// Pick the active theme in:
//   Tools > Options > PresentationAssistant > General > Appearance > Theme
// Changes here show up on the next keystroke; no restart needed.
//
// Built-in names: {builtIns}
//
// Their exact colours are listed in themes.reference.json, next to this file -
// copy an entry from there and tweak it. ""Auto"" follows Visual Studio, using
// Classic and ClassicDark, so overriding either changes what Auto looks like.
//
// Example - remove the surrounding /* */ to switch it on:
/*
[
  {{
    ""Name"": ""My Talk"",
    ""Background"": ""#101820"",
    ""Foreground"": ""#F2F2F2"",
    ""SecondaryForeground"": ""#8AA0B0"",
    ""Border"": ""#2E86AB"",
    ""Opacity"": 0.9,
    ""IsDark"": true
  }},
  {{
    ""Name"": ""Amber"",
    ""Opacity"": 1.0
  }}
]
*/
[]
";
        }

        private static ThemePalette P(
            string name, string background, string foreground,
            string secondary, string border, double opacity, bool isDark)
        {
            return new ThemePalette(name, Hex(background), Hex(foreground),
                Hex(secondary), Hex(border), opacity, isDark);
        }

        private static Color Hex(string value) => (Color)ColorConverter.ConvertFromString(value);
    }
}
