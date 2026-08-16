# PresentationAssistant

A Visual Studio extension that announces the command you just ran, and the keyboard
shortcut that runs it, in a small overlay just above the status bar.

Useful when you are screen sharing, recording a demo, pairing, teaching — or just trying
to learn your own key bindings.

![image](https://github.com/lokeshgovindu/PresentationAssistant/assets/8664883/376f7106-537d-4131-b829-3bae911cc4dd)

![image](https://github.com/lokeshgovindu/PresentationAssistant/assets/8664883/2c05e86b-fe21-4b1e-9784-bdbda4eebbce)

## Contents

- [What you see](#what-you-see)
- [Installing](#installing)
- [Requirements](#requirements)
- [Options](#options)
- [Excluding commands](#excluding-commands)
- [Themes](#themes)
- [Custom themes](#custom-themes)
- [Where things are stored](#where-things-are-stored)
- [Building from source](#building-from-source)
- [How it works](#how-it-works)
- [Project layout](#project-layout)

## What you see

![The overlay](docs/overlay.png)

```
Scroll Line Down via Ctrl+Down Arrow ×9
└──────┬──────┘     └──────┬──────┘  └┬┘
   command          key binding    repeat count
```

- **Command** — the command's display name, in bold. The full command id goes to the
  status bar as `PA: Edit.ScrollLineDown via Ctrl+Down Arrow`, which is also the form to
  use in [Excluded Commands](#excluding-commands).
- **Key binding** — every shortcut bound to the command. If there is more than one they
  are joined with `or`. Commands with no binding still appear, unless you turn on
  [Shortcuts Only](#options).
- **Repeat count** — holding a key down, or pressing the same shortcut repeatedly, does
  not flash the overlay once per keystroke. Consecutive invocations of the same command
  collapse into one line with a `×N` counter. A run ends when a different command is
  invoked, or when **Multiplier Timeout** elapses between two keystrokes.

The overlay never takes focus, never appears in Alt+Tab or the taskbar, and hides itself
after **Window Timeout**.

## Installing

From the [Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=LokeshGovindu.PresentationAssistant),
or by double-clicking a `PresentationAssistant.vsix` built from source.

There is nothing to switch on afterwards — the extension loads with the IDE and starts
announcing commands immediately.

## Requirements

| | |
| --- | --- |
| Visual Studio | 2022 (17.x) or 2026 (18.x), `amd64` |
| .NET Framework | 4.7.2 |

Both are supported by the same VSIX — the manifest's `InstallationTarget` is
`[17.0, 19.0)`.

## Options

**Tools → Options → PresentationAssistant → General**

![The options page](docs/options.png)

| Setting | Default | Meaning |
| --- | --- | --- |
| **Window Timeout (MS)** | `5000` | How long the overlay stays on screen after the last command. |
| **Multiplier Timeout (MS)** | `10000` | Maximum gap between two presses that still count towards the same `×N` run. |
| **Shortcuts Only** | `false` | When set, commands without a key binding are not announced. |
| **Excluded Commands** | empty | Commands never to announce — see [below](#excluding-commands). |
| **Theme** | `VisualStudio` | Overlay colours — see [below](#themes). |
| **Themes File** | — | Shows where `themes.json` lives. Press the **`...`** button to open it for editing. |

Changes apply immediately; there is no need to restart the IDE.

## Excluding commands

Some commands are just noise — whatever you happen to hold down, or whatever your own
setup fires constantly. List them in **Excluded Commands**, separated by semicolons:

```
Edit.Line*; View.Output; File.SaveAll
```

- Patterns match the command id shown in the status bar, e.g. `Edit.ScrollLineDown`.
- A trailing `*` matches a prefix, so `Edit.Line*` covers `Edit.LineDown`, `Edit.LineUp`
  and friends.
- Matching is case-insensitive, and `;`, `,` and newlines all separate patterns.
- A bare `*` is ignored rather than silencing everything.

In the settings file this is an array, which is easier to keep tidy for a long list:

```jsonc
"ExcludedCommands": [ "Edit.Line*", "View.Output" ]
```

A handful of commands the IDE fires on its own are **always** excluded, so you never have
to name them: cursor and page movement (`Edit.LineUp`, `Edit.CharLeft`, `Edit.PageDown`, …),
the debugger's location toolbar combo boxes (`Debug.LocationToolbar.*`), and the build
configuration and platform dropdowns.

## Themes

Two themes adapt to the IDE:

| Theme | Look |
| --- | --- |
| `VisualStudio` | **Default.** Matches the running IDE — built from the shell's own tool window colours, tinted towards the accent so the overlay reads as a callout instead of disappearing into the chrome. |
| `Auto` | Follows the IDE more loosely: the `Classic` palette under the Blue and Light themes, `ClassicDark` under Dark. Pick this for the extension's original green look while still tracking light/dark. |

The rest are fixed palettes — eight for light shells, eleven for dark:

![The built-in themes](docs/themes.png)

`Classic` is the pale green overlay the extension originally shipped with, and
`ClassicDark` is its green-tinted counterpart. The earlier names `Light` and `Dark` still
resolve to them, so an existing setting keeps working.

Both adaptive themes re-resolve when you switch the Visual Studio theme, so the overlay
never stays bright green over a dark shell. Neither depends on a theme GUID:
`VisualStudio` reads live shell colours, and `Auto` decides light from dark by the
brightness of the tool window background — so custom and third-party themes are handled
correctly too. `VisualStudio` also checks contrast and substitutes black or white text if
a theme would otherwise render the overlay unreadable.

## Custom themes

Themes live in `%APPDATA%\PresentationAssistant\themes.json`.

The quickest way in is the **`...`** button on the **Themes File** row in Options — it
creates the file if needed and opens it in the editor. Alongside it sits
`themes.reference.json`, a generated listing of every built-in theme in full, so editing
one is copy, paste, tweak rather than guessing hex codes. That file is rewritten every
session and never read back; only `themes.json` needs your attention.

An entry adds a new theme, or overrides a built-in one when the name matches:

```jsonc
[
  {
    "Name": "My Talk",
    "Background": "#101820",
    "Foreground": "#F2F2F2",
    "SecondaryForeground": "#8AA0B0",
    "Border": "#2E86AB",
    "Opacity": 0.9,
    "IsDark": true
  },

  // Keep Amber's colours, just make it fully opaque.
  { "Name": "Amber", "Opacity": 1.0 }
]
```

### Fields

| Field | Meaning |
| --- | --- |
| `Name` | **Required.** The name shown in the Theme dropdown. Matching a built-in overrides it. |
| `Background` | Overlay background. |
| `Foreground` | The bold command name. |
| `SecondaryForeground` | The dimmer `via …` and `×N` runs. Omit it and it is derived by fading `Foreground` towards `Background` — usually the right answer, so most themes only need three colours. |
| `Border` | The one-pixel border. |
| `Opacity` | `0.1`–`1.0`. Values outside that range are clamped. |
| `IsDark` | Marks the theme as belonging to a dark shell, which is what makes it eligible for the `Auto` pairing. Omit it and it is inferred from the background brightness. |

Everything except `Name` is optional: what you leave out is inherited from the built-in of
the same name, or from `Classic` for a brand new theme.

Colours accept anything WPF understands — `#RRGGBB`, `#AARRGGBB`, or a named colour such
as `Gainsboro`. Comments are allowed anywhere in the file.

### Behaviour

- Your themes appear in the Theme dropdown alongside the built-ins.
- Edits apply on the next keystroke; no restart needed.
- A malformed file, an unparseable colour, or an entry with no `Name` falls back to the
  built-ins rather than breaking the overlay.
- The extension only ever *reads* `themes.json` after creating it, so your formatting and
  comments are preserved.

### Tuning the adaptive themes

Overriding `Classic` or `ClassicDark` also changes what `Auto` looks like — the neatest way
to customise the overlay while still following the IDE:

```jsonc
[
  { "Name": "Classic",     "Background": "#EAF7EA", "Border": "#4C9A4C" },
  { "Name": "ClassicDark",  "Background": "#12211A", "Border": "#4C9A4C" }
]
```

`Auto` and `VisualStudio` are computed rather than stored, but they can be tuned the same
way: an entry named after either is merged on top of the palette it worked out, so you can
nudge just the part you want.

```jsonc
[
  // Keep matching the IDE, but make it fully opaque with a green border.
  { "Name": "VisualStudio", "Opacity": 1.0, "Border": "#4C9A63" }
]
```

## Where things are stored

All under `%APPDATA%\PresentationAssistant\`:

| File | Purpose |
| --- | --- |
| `presentationassistant.json` | Your settings. Written by the Options page, and safe to edit by hand. |
| `themes.json` | Your custom themes. Created once, then only read. |
| `themes.reference.json` | Generated listing of the built-in themes, to copy from. Rewritten each session; editing it has no effect. |

Because settings live here rather than in the Visual Studio registry hive, they survive
IDE upgrades and can be copied between machines.

## Building from source

```powershell
msbuild PresentationAssistant.csproj /t:Restore
msbuild PresentationAssistant.csproj /p:Configuration=Release /p:Platform=AnyCPU
```

The VSIX is written to `bin\Release\PresentationAssistant.vsix`.

`F5` launches an experimental instance (`devenv.exe /rootsuffix Exp`) with the extension
deployed, which is the easiest way to try a change.

> **Which MSBuild you use decides which experimental hive gets the extension.** VS 2022's
> MSBuild deploys to `17.0Exp`, VS 2026's to `18.0Exp`; both can load it. To target a
> specific one — handy when `VisualStudioVersion` is already set in your environment —
> pass it explicitly:
>
> ```powershell
> # deploy to the Visual Studio 2022 experimental instance
> & "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe" `
>   PresentationAssistant.csproj /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU `
>   /p:VisualStudioVersion=17.0 /p:DeployExtension=true
> ```
>
> Run `/t:Restore` again after switching, since the restore is toolset-specific. Close any
> running experimental instance first, or the deploy fails on a locked DLL.

Newtonsoft.Json is a compile-time-only reference — it is not copied into the VSIX, and
binds to the copy the shell already loads.

## How it works

The package hooks `DTE.Events.CommandEvents.BeforeExecute`, which the IDE raises for every
command it is about to run, however that command was invoked — key press, menu, toolbar or
automation. For each one it looks the command up to get its display name and key bindings,
drops it if a blocklist matches, updates the repeat counter, and shows a borderless WPF
window positioned just above the status bar.

That hook is why the overlay reports commands rather than raw keystrokes: it shows what
the IDE actually did, so a shortcut that is unbound, overridden, or swallowed by another
extension is reported accurately.

The overlay window is created once and reused, which costs roughly 3ms per command instead
of 17ms to build a new one. Between commands it is parked off-screen rather than hidden:
WPF stops composing a hidden window, so showing one again briefly presents the last frame
it drew — the previous command. Parking keeps it composing, so a new command can update
the content out of sight and only move into view once that content has actually been
rendered.

## Project layout

| Path | Role |
| --- | --- |
| `PresentationAssistantPackage.cs` | `AsyncPackage`; the command hook, filtering, and overlay lifetime. |
| `PresentationAssistantWindow.xaml[.cs]` | The overlay window — data binding, placement, hide timer, theme application. |
| `ShortcutDetails.cs` | View model bound to the overlay. |
| `ShortcutDisplayStatistics.cs` | The `×N` repeat counter. |
| `ActionIdBlocklist.cs` | Built-in list of commands never announced. |
| `CommandExclusions.cs` | The user's **Excluded Commands** patterns. |
| `AppPaths.cs` | Locations of the settings and theme files. |
| `SampleData.cs` | Design-time data for the XAML designer. |
| `State\Settings.cs` | Settings model and JSON persistence. |
| `State\PresentationAssistantOptionsDialog.cs` | The Tools → Options page. |
| `Theming\ThemeNames.cs` | The names with behaviour attached (`Auto`, `VisualStudio`, `Classic`, `ClassicDark`). |
| `Theming\ThemePalette.cs` | A resolved palette, plus the colour maths. |
| `Theming\ThemeDefinition.cs` | One `themes.json` entry, and how it merges onto a baseline. |
| `Theming\ThemeCatalog.cs` | Built-in palettes, `themes.json` loading, and the generated reference file. |
| `Theming\ThemeManager.cs` | Resolves a name to a palette; follows the shell theme. |
| `Theming\ThemeNameConverter.cs` | Populates the Theme dropdown from the catalog. |
| `Theming\ThemesFileEditor.cs` | The **`...`** button that opens `themes.json`. |
