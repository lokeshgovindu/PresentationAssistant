Announces the command you just ran, and the keyboard shortcut that runs it, in a small overlay just above the status bar.

Useful when you are screen sharing, recording a demo, pairing, teaching - or just trying to learn your own key bindings.

![PA-1.png](PA-1.png)

## What you see

![overlay.png](overlay.png)

- The **command name**, in bold. The full command id also goes to the status bar, as `PA: Edit.ScrollLineDown via Ctrl+Down Arrow`.
- Every **key binding** for that command. If there is more than one they are joined with `or`.
- A **repeat count**. Holding a key down, or pressing the same shortcut repeatedly, shows a single line with `x9` rather than flashing nine times. A run ends when you invoke a different command, or after the multiplier timeout.

The overlay never takes focus, never appears in Alt+Tab or the taskbar, is transparent to the mouse so it never swallows a click, and hides itself after a few seconds.

## Themes

The default theme, **VisualStudio**, builds the overlay from the running IDE's own colours - tinted towards your accent so it still reads as a callout instead of disappearing into the chrome - and re-resolves itself whenever you switch theme.

Nineteen fixed palettes come with it, including the original green overlay as **Classic** and its dark counterpart **ClassicDark**, plus Nord, Dracula, Monokai, Gruvbox Dark, One Dark, Tokyo Night and both Solarizeds. Choosing **Auto** gives you the Classic / ClassicDark pair, switching with the IDE.

![themes.png](themes.png)

## Your own themes

Themes live in `%APPDATA%\PresentationAssistant\themes.json`. An entry adds a theme, or overrides a built-in one when the name matches - and every field except `Name` is optional, so you can change just the one colour you care about:

    [
      { "Name": "My Talk", "Background": "#101820", "Border": "#2E86AB", "IsDark": true },
      { "Name": "Amber", "Opacity": 1.0 }
    ]

Press the `...` button on the **Themes File** row in Options to open it. A generated `themes.reference.json` sits beside it listing every built-in theme in full, so editing one is copy, paste, tweak. Changes apply on the next keystroke, with no restart.

Overriding `Classic` or `ClassicDark` also changes what `Auto` looks like.

## Options

Under **Tools > Options > PresentationAssistant > General**:

![options.png](options.png)

- **Theme** - overlay colours, including any you define yourself.
- **Themes File** - opens `themes.json` for editing.
- **Window Timeout (MS)** - how long the overlay stays on screen. Default 5000.
- **Multiplier Timeout (MS)** - the largest gap between two presses that still count as a repeat. Default 10000.
- **Shortcuts Only** - when set, commands without a key binding are not announced.
- **Excluded Commands** - commands never to announce.

Settings are stored as JSON in `%APPDATA%\PresentationAssistant`, so they survive Visual Studio upgrades and can be copied between machines.

## Excluding commands

Some commands are just noise. List them in **Excluded Commands**, separated by semicolons, with a trailing `*` to match a prefix:

    Edit.Line*; View.Output; File.SaveAll

Patterns match the command id shown in the status bar, and matching is case-insensitive. A set of commands the IDE fires constantly on its own - cursor movement, the debugger's location toolbar, the build configuration dropdowns - is always excluded on top of your list.

## Requirements

Visual Studio 2022 or 2026, 64-bit.

## Links

- [Source code](https://github.com/lokeshgovindu/PresentationAssistant)
- [Release notes](https://github.com/lokeshgovindu/PresentationAssistant/releases)
- [Report an issue](https://github.com/lokeshgovindu/PresentationAssistant/issues)
