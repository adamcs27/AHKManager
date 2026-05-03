# AHK Script Manager
.NET Framework WinForms application for managing AutoHotkey scripts.

## Features

- **Drag & Drop** `.ahk` files directly into the window, or use the Browse button
- **Groups / Tabs** — create named groups and assign scripts to them via right-click
- **Run / Reload / Edit** — per-script action buttons; Edit opens in Notepad
- **Running Status Indicator** — green when active, grey when stopped
- **Hotkey Detection** — automatically parses your `.ahk` files for hotkey bindings
- **Script Descriptions** — right-click → Add/Edit Description
- **Search Bar** — filter scripts by name or description
- **Run All / Reload All** — within scope of the current group tab
- **Global Suspend Hotkey** — configurable in Settings 
- **Dark Theme** — native Windows look with custom dark UI

## Building

### Requirements
- Windows 10/11
- Visual Studio 2019 or 2022 (any edition, including free Community)
  - OR: [Build Tools for Visual Studio](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022)
- .NET Framework 4.7.2 (included with Windows 10+)

### Steps

**Build script:**
1. Double-click `Build.bat`
2. The exe will be in `bin\Release\AHKScriptManager.exe`

## Usage

### Adding Scripts
- **Drag and drop** `.ahk` files onto the window
- Click **Browse...** to open a file picker

### Groups
- Click **+ Group** in the toolbar to create a named group
- Right-click any script → **Send to Group** to assign it
- Switch between groups using the tabs at the top

### Running Scripts
- Click **Run** to launch a script (**Stop** appears when running)
- Click **Reload** to restart a script
- Click **Edit** to open the script in Notepad

### Settings
- Click Gear icon → set a global suspend hotkey (e.g. `^!s` for Ctrl+Alt+S)
- AHK modifier syntax: `^` = Ctrl, `!` = Alt, `+` = Shift, `#` = Win

## Data Storage
Settings and script list are saved to:
```
%AppData%\AHKScriptManager\data.xml
```

## Notes
- AutoHotkey must be installed for scripts to run (or `.ahk` files must be associated with AHK)
- The app checks common AHK installation paths automatically
- Process status is polled every 1.5 seconds to detect when scripts exit, stops when minimized.
