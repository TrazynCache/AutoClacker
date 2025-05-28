# AutoClacker

## Version
1.0.0.4 (as of May 28, 2025)

AutoClacker is a Windows application that automates mouse clicks and keyboard key presses. It supports global or application-restricted modes, making it suitable for offline tasks like automating actions in games or applications.

![AutoClacker Screenshot](../Images/AutoClacker_Not_Running.png)

## Features

- **Global or Restricted Mode**: Automate actions system-wide or limit them to a specific application.
- **Mouse and Keyboard Automation**: Configure mouse clicks (single/double, click/hold) or keyboard key presses (press/hold).
- **Customizable Settings**: Adjust intervals, durations, and hotkeys via a user-friendly GUI.
- **Persistent Settings**: Settings are saved in a JSON file (`settings.json`) for consistency across sessions.
- **Hotkey Support**: Toggle automation with a configurable hotkey (default: F5).
- **Theme Support**: Switch between light and dark modes for the user interface.

## Installation

1. **Using the Released Executable**:
   - Download the latest release package containing `AutoClacker.exe` and `AutoClacker.exe.config` (located in `bin\Release` after building).
   - Run `AutoClacker.exe` directly. A `settings.json` file will be generated in the same directory on first run or if deleted, containing user settings (e.g., theme, hotkeys).

2. **Building from Source**:
   - Clone or download the repository to your local machine.
   - Open the solution (`AutoClacker.sln`) in Visual Studio 2022 or a compatible C# IDE.
   - Build the project to resolve dependencies.
   - Run the application from the IDE or the compiled executable (`bin\Debug\AutoClacker.exe` for debugging or `bin\Release\AutoClacker.exe` for release).

## Usage

1. **Select Scope**: Choose "Global" or "Restricted to Application" mode.
2. **Configure Actions**: Set mouse (button, click type, mode) or keyboard (key, mode) actions.
3. **Set Toggle Key**: Define a hotkey to start/stop automation (default: F5).
4. **Adjust Speed and Duration**: Customize the interval between actions and total duration (if applicable).
5. **Start/Stop Automation**: Press the toggle key to begin or end automation.
6. **Change Theme**: Use the Options dialog to switch between light and dark modes.

## Project Structure

- `App.xaml`: Application entry point and global resource definitions (e.g., converters).
- `Controllers/AutomationController.cs`: Core logic for mouse and keyboard automation.
- `Converters/*.cs`: Data-binding converters for WPF UI (e.g., `StringToBooleanConverter`, `BooleanToVisibilityConverter`).
- `Docs/README.markdown`: Project documentation (this file).
- `Docs/LICENSE.txt`: MIT License details.
- `Images/AUTOCLACKER_NOT_RUNNING.png`: Screenshot of the application (unused in runtime).
- `Models/Settings.cs`: Defines configuration settings for serialization to `settings.json`.
- `Properties/AssemblyInfo.cs`: Assembly metadata and version information.
- `Properties/Resources.resx`: Resource file for localized strings.
- `Themes/LightTheme.xaml`: Styles for the light theme.
- `Themes/DarkTheme.xaml`: Styles for the dark theme.
- `Utilities/ApplicationDetector.cs`: Detects running applications for restricted mode.
- `Utilities/HotkeyManager.cs`: Handles global hotkeys.
- `Utilities/SettingsManager.cs`: Manages loading and saving settings to `settings.json`.
- `ViewModels/MainViewModel.cs`: Coordinates UI and automation logic.
- `Views/MainWindow.xaml`: Main GUI for configuring and controlling automation.
- `Views/OptionsDialog.xaml`: Dialog for theme selection (light/dark mode) and other settings.

## License

Licensed under the MIT License. See `Docs/LICENSE.txt` for details.