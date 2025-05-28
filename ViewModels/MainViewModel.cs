using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using AutoClacker.Models;
using AutoClacker.Controllers;
using AutoClacker.Utilities;
using AutoClacker.Views;
using System.Threading.Tasks;

namespace AutoClacker.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly Settings settings;
        private readonly AutomationController automationController;
        private readonly ApplicationDetector applicationDetector;
        private HotkeyManager hotkeyManager;
        private bool isRunning;
        private string statusText = "Not running";
        private string statusColor = "Red";
        private Key capturedKey;
        private bool isSettingToggleKey;
        private bool isSettingKeyboardKey;
        private List<string> runningApplications;
        private List<string> mouseButtonOptions;
        private List<string> clickTypeOptions;
        private Window window;
        private TaskCompletionSource<bool> automationTcs;
        private readonly DispatcherTimer uiUpdateTimer;
        private TimeSpan remainingDuration;
        private int remainingDurationMin;
        private int remainingDurationSec;
        private int remainingDurationMs;
        private DateTime? durationStartTime;

        public event PropertyChangedEventHandler PropertyChanged;

        public MainViewModel()
        {
            settings = SettingsManager.LoadSettings();
            automationController = new AutomationController(this);
            applicationDetector = new ApplicationDetector();
            RunningApplications = applicationDetector.GetRunningApplications();
            MouseButtonOptions = new List<string> { "Left", "Right", "Middle" };
            ClickTypeOptions = new List<string> { "Single", "Double" };
            ToggleAutomationCommand = new RelayCommand(async o => await ToggleAutomation());
            SetTriggerKeyCommand = new RelayCommand(async o => await SetTriggerKey());
            SetKeyCommand = new RelayCommand(async o => await SetKey());
            ResetSettingsCommand = new RelayCommand(async o => await ResetSettings());
            SetConstantCommand = new RelayCommand(SetConstant);
            SetHoldDurationCommand = new RelayCommand(SetHoldDuration);
            RefreshApplicationsCommand = new RelayCommand(RefreshApplications);
            OpenOptionsCommand = new RelayCommand(async o => await OpenOptionsDialog());

            uiUpdateTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            uiUpdateTimer.Tick += (s, e) => UpdateRemainingDurationDisplay();

            OnPropertyChanged(nameof(TriggerKeyDisplay));
            OnPropertyChanged(nameof(KeyboardKeyDisplay));
            OnPropertyChanged(nameof(IsMouseMode));
            OnPropertyChanged(nameof(IsKeyboardMode));
            OnPropertyChanged(nameof(IsClickModeVisible));
            OnPropertyChanged(nameof(IsHoldModeVisible));
            OnPropertyChanged(nameof(IsClickDurationMode));
            OnPropertyChanged(nameof(IsHoldDurationMode));
            OnPropertyChanged(nameof(IsPressModeVisible));
            OnPropertyChanged(nameof(IsHoldModeVisibleKeyboard));
            OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
            OnPropertyChanged(nameof(IsTimerMode));
            OnPropertyChanged(nameof(IsRestrictedMode));
            Console.WriteLine($"Initial MouseMode: {MouseMode}, ClickMode: {ClickMode}, HoldMode: {HoldMode}, KeyboardMode: {KeyboardMode}");
        }

        public MainViewModel(Window window) : this()
        {
            this.window = window;
        }

        public void InitializeHotkeyManager(Window window)
        {
            this.window = window;
            hotkeyManager = new HotkeyManager(window, this);
            hotkeyManager.RegisterTriggerHotkey(settings.TriggerKey, settings.TriggerKeyModifiers);
            window.Topmost = settings.IsTopmost;
        }

        public TimeSpan GetRemainingDuration()
        {
            return remainingDuration;
        }

        private void UpdateRemainingDurationDisplay()
        {
            if (durationStartTime.HasValue)
            {
                var elapsed = DateTime.Now - durationStartTime.Value;
                remainingDuration = GetActiveDuration() - elapsed;
                if (remainingDuration <= TimeSpan.Zero)
                {
                    remainingDuration = TimeSpan.Zero;
                    _ = ManageAutomationAsync(false, "Duration completed");
                }

                RemainingDurationMin = remainingDuration.Minutes;
                RemainingDurationSec = remainingDuration.Seconds;
                RemainingDurationMs = remainingDuration.Milliseconds;

                OnPropertyChanged(nameof(RemainingDurationMin));
                OnPropertyChanged(nameof(RemainingDurationSec));
                OnPropertyChanged(nameof(RemainingDurationMs));
            }
        }

        private TimeSpan GetActiveDuration()
        {
            if (settings.ActionType == "Mouse" && settings.MouseMode == "Click" && settings.ClickMode == "Duration")
            {
                return settings.ClickDuration;
            }
            if (settings.ActionType == "Mouse" && settings.MouseMode == "Hold" && settings.HoldMode == "HoldDuration")
            {
                return settings.MouseHoldDuration;
            }
            if (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Hold" && settings.KeyboardHoldDuration != TimeSpan.Zero)
            {
                return settings.KeyboardHoldDuration;
            }
            if (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Press" && settings.Mode == "Timer")
            {
                return settings.TotalDuration;
            }
            return TimeSpan.Zero;
        }

        public int RemainingDurationMin
        {
            get => remainingDurationMin;
            private set { remainingDurationMin = value; OnPropertyChanged(nameof(RemainingDurationMin)); }
        }

        public int RemainingDurationSec
        {
            get => remainingDurationSec;
            private set { remainingDurationSec = value; OnPropertyChanged(nameof(RemainingDurationSec)); }
        }

        public int RemainingDurationMs
        {
            get => remainingDurationMs;
            private set { remainingDurationMs = value; OnPropertyChanged(nameof(RemainingDurationMs)); }
        }

        public void StartTimers()
        {
            Console.WriteLine("StartTimers called.");
            if (isRunning)
            {
                Console.WriteLine("Automation already running, skipping StartTimers.");
                return;
            }

            if ((settings.ActionType == "Mouse" && settings.MouseMode == "Click" && settings.ClickMode == "Duration") ||
                (settings.ActionType == "Mouse" && settings.MouseMode == "Hold" && settings.HoldMode == "HoldDuration") ||
                (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Hold" && settings.KeyboardHoldDuration != TimeSpan.Zero) ||
                (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Press" && settings.Mode == "Timer"))
            {
                remainingDuration = GetActiveDuration();
                durationStartTime = DateTime.Now;
                uiUpdateTimer.Start();
                Console.WriteLine("UI update timer started.");
            }
            else
            {
                Console.WriteLine("Constant mode detected, skipping timer initialization.");
            }
        }

        public void StopTimers()
        {
            Console.WriteLine("StopTimers called.");
            uiUpdateTimer.Stop();
            remainingDuration = TimeSpan.Zero;
            durationStartTime = null;
            RemainingDurationMin = 0;
            RemainingDurationSec = 0;
            RemainingDurationMs = 0;
        }

        public void UpdateStatus(string text, string color)
        {
            StatusText = text;
            StatusColor = color;
            isRunning = text == "Running";
            OnPropertyChanged(nameof(IsRunning));
        }

        public Settings CurrentSettings => settings;

        public string ClickScope
        {
            get => settings.ClickScope;
            set
            {
                settings.ClickScope = value;
                OnPropertyChanged(nameof(ClickScope));
                OnPropertyChanged(nameof(IsRestrictedMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string TargetApplication
        {
            get => settings.TargetApplication;
            set
            {
                settings.TargetApplication = value;
                OnPropertyChanged(nameof(TargetApplication));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string ActionType
        {
            get => settings.ActionType;
            set
            {
                settings.ActionType = value;
                OnPropertyChanged(nameof(ActionType));
                OnPropertyChanged(nameof(IsMouseMode));
                OnPropertyChanged(nameof(IsKeyboardMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string MouseButton
        {
            get => settings.MouseButton;
            set
            {
                settings.MouseButton = value;
                OnPropertyChanged(nameof(MouseButton));
                SettingsManager.SaveSettings(settings);
            }
        }

        public List<string> MouseButtonOptions
        {
            get => mouseButtonOptions;
            private set { mouseButtonOptions = value; OnPropertyChanged(nameof(MouseButtonOptions)); }
        }

        public string ClickType
        {
            get => settings.ClickType;
            set
            {
                settings.ClickType = value;
                OnPropertyChanged(nameof(ClickType));
                SettingsManager.SaveSettings(settings);
            }
        }

        public List<string> ClickTypeOptions
        {
            get => clickTypeOptions;
            private set { clickTypeOptions = value; OnPropertyChanged(nameof(ClickTypeOptions)); }
        }

        public string MouseMode
        {
            get => settings.MouseMode;
            set
            {
                settings.MouseMode = value;
                OnPropertyChanged(nameof(MouseMode));
                OnPropertyChanged(nameof(IsClickModeVisible));
                OnPropertyChanged(nameof(IsHoldModeVisible));
                OnPropertyChanged(nameof(IsClickDurationMode));
                OnPropertyChanged(nameof(IsHoldDurationMode));
                SettingsManager.SaveSettings(settings);
                Console.WriteLine($"MouseMode changed to: {value}, IsClickModeVisible: {IsClickModeVisible}, IsHoldModeVisible: {IsHoldModeVisible}");
            }
        }

        public string ClickMode
        {
            get => settings.ClickMode;
            set
            {
                settings.ClickMode = value;
                OnPropertyChanged(nameof(ClickMode));
                OnPropertyChanged(nameof(IsClickDurationMode));
                SettingsManager.SaveSettings(settings);
                Console.WriteLine($"ClickMode changed to: {value}, IsClickDurationMode: {IsClickDurationMode}");
            }
        }

        public int ClickDurationMinutes
        {
            get => settings.ClickDuration.Minutes;
            set
            {
                settings.ClickDuration = new TimeSpan(0, 0, value, settings.ClickDuration.Seconds, settings.ClickDuration.Milliseconds);
                OnPropertyChanged(nameof(ClickDurationMinutes));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int ClickDurationSeconds
        {
            get => settings.ClickDuration.Seconds;
            set
            {
                settings.ClickDuration = new TimeSpan(0, 0, settings.ClickDuration.Minutes, value, settings.ClickDuration.Milliseconds);
                OnPropertyChanged(nameof(ClickDurationSeconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int ClickDurationMilliseconds
        {
            get => settings.ClickDuration.Milliseconds;
            set
            {
                settings.ClickDuration = new TimeSpan(0, 0, settings.ClickDuration.Minutes, settings.ClickDuration.Seconds, value);
                OnPropertyChanged(nameof(ClickDurationMilliseconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int MouseHoldDurationMinutes
        {
            get => settings.MouseHoldDuration.Minutes;
            set
            {
                settings.MouseHoldDuration = new TimeSpan(0, 0, value, settings.MouseHoldDuration.Seconds, settings.MouseHoldDuration.Milliseconds);
                OnPropertyChanged(nameof(MouseHoldDurationMinutes));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int MouseHoldDurationSeconds
        {
            get => settings.MouseHoldDuration.Seconds;
            set
            {
                settings.MouseHoldDuration = new TimeSpan(0, 0, settings.MouseHoldDuration.Minutes, value, settings.MouseHoldDuration.Milliseconds);
                OnPropertyChanged(nameof(MouseHoldDurationSeconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int MouseHoldDurationMilliseconds
        {
            get => settings.MouseHoldDuration.Milliseconds;
            set
            {
                settings.MouseHoldDuration = new TimeSpan(0, 0, settings.MouseHoldDuration.Minutes, settings.MouseHoldDuration.Seconds, value);
                OnPropertyChanged(nameof(MouseHoldDurationMilliseconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string HoldMode
        {
            get => settings.HoldMode;
            set
            {
                settings.HoldMode = value;
                OnPropertyChanged(nameof(HoldMode));
                OnPropertyChanged(nameof(IsHoldDurationMode));
                SettingsManager.SaveSettings(settings);
                Console.WriteLine($"HoldMode changed to: {value}, IsHoldDurationMode: {IsHoldDurationMode}");
            }
        }

        public bool MousePhysicalHoldMode
        {
            get => settings.MousePhysicalHoldMode;
            set
            {
                settings.MousePhysicalHoldMode = value;
                OnPropertyChanged(nameof(MousePhysicalHoldMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public Key KeyboardKey
        {
            get => settings.KeyboardKey;
            set
            {
                settings.KeyboardKey = value;
                OnPropertyChanged(nameof(KeyboardKey));
                OnPropertyChanged(nameof(KeyboardKeyDisplay));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string KeyboardKeyDisplay => settings.KeyboardKey.ToString();

        public string KeyboardMode
        {
            get => settings.KeyboardMode;
            set
            {
                settings.KeyboardMode = value;
                OnPropertyChanged(nameof(KeyboardMode));
                OnPropertyChanged(nameof(IsPressModeVisible));
                OnPropertyChanged(nameof(IsHoldModeVisibleKeyboard));
                OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
                SettingsManager.SaveSettings(settings);
                Console.WriteLine($"KeyboardMode changed to: {value}, IsPressModeVisible: {IsPressModeVisible}, IsHoldModeVisibleKeyboard: {IsHoldModeVisibleKeyboard}");
            }
        }

        public int KeyboardHoldDurationMinutes
        {
            get => settings.KeyboardHoldDuration.Minutes;
            set
            {
                settings.KeyboardHoldDuration = new TimeSpan(0, 0, value, settings.KeyboardHoldDuration.Seconds, settings.KeyboardHoldDuration.Milliseconds);
                OnPropertyChanged(nameof(KeyboardHoldDurationMinutes));
                OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int KeyboardHoldDurationSeconds
        {
            get => settings.KeyboardHoldDuration.Seconds;
            set
            {
                settings.KeyboardHoldDuration = new TimeSpan(0, 0, settings.KeyboardHoldDuration.Minutes, value, settings.KeyboardHoldDuration.Milliseconds);
                OnPropertyChanged(nameof(KeyboardHoldDurationSeconds));
                OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int KeyboardHoldDurationMilliseconds
        {
            get => settings.KeyboardHoldDuration.Milliseconds;
            set
            {
                settings.KeyboardHoldDuration = new TimeSpan(0, 0, settings.KeyboardHoldDuration.Minutes, settings.KeyboardHoldDuration.Seconds, value);
                OnPropertyChanged(nameof(KeyboardHoldDurationMilliseconds));
                OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public bool KeyboardPhysicalHoldMode
        {
            get => settings.KeyboardPhysicalHoldMode;
            set
            {
                settings.KeyboardPhysicalHoldMode = value;
                OnPropertyChanged(nameof(KeyboardPhysicalHoldMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public Key TriggerKey
        {
            get => settings.TriggerKey;
            set
            {
                settings.TriggerKey = value;
                OnPropertyChanged(nameof(TriggerKey));
                OnPropertyChanged(nameof(TriggerKeyDisplay));
                SettingsManager.SaveSettings(settings);
                hotkeyManager?.RegisterTriggerHotkey(value, settings.TriggerKeyModifiers);
            }
        }

        public string TriggerKeyDisplay => settings.TriggerKey.ToString();

        public ModifierKeys TriggerKeyModifiers
        {
            get => settings.TriggerKeyModifiers;
            set
            {
                settings.TriggerKeyModifiers = value;
                OnPropertyChanged(nameof(TriggerKeyModifiers));
                SettingsManager.SaveSettings(settings);
                hotkeyManager?.RegisterTriggerHotkey(settings.TriggerKey, value);
            }
        }

        public int IntervalMinutes
        {
            get => settings.Interval.Minutes;
            set
            {
                settings.Interval = new TimeSpan(0, 0, value, settings.Interval.Seconds, settings.Interval.Milliseconds);
                OnPropertyChanged(nameof(IntervalMinutes));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int IntervalSeconds
        {
            get => settings.Interval.Seconds;
            set
            {
                settings.Interval = new TimeSpan(0, 0, settings.Interval.Minutes, value, settings.Interval.Milliseconds);
                OnPropertyChanged(nameof(IntervalSeconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int IntervalMilliseconds
        {
            get => settings.Interval.Milliseconds;
            set
            {
                settings.Interval = new TimeSpan(0, 0, settings.Interval.Minutes, settings.Interval.Seconds, value);
                OnPropertyChanged(nameof(IntervalMilliseconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string Mode
        {
            get => settings.Mode;
            set
            {
                settings.Mode = value;
                OnPropertyChanged(nameof(Mode));
                OnPropertyChanged(nameof(IsTimerMode));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int TotalDurationMinutes
        {
            get => settings.TotalDuration.Minutes;
            set
            {
                settings.TotalDuration = new TimeSpan(0, 0, value, settings.TotalDuration.Seconds, settings.TotalDuration.Milliseconds);
                OnPropertyChanged(nameof(TotalDurationMinutes));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int TotalDurationSeconds
        {
            get => settings.TotalDuration.Seconds;
            set
            {
                settings.TotalDuration = new TimeSpan(0, 0, settings.TotalDuration.Minutes, value, settings.TotalDuration.Milliseconds);
                OnPropertyChanged(nameof(TotalDurationSeconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public int TotalDurationMilliseconds
        {
            get => settings.TotalDuration.Milliseconds;
            set
            {
                settings.TotalDuration = new TimeSpan(0, 0, settings.TotalDuration.Minutes, settings.TotalDuration.Seconds, value);
                OnPropertyChanged(nameof(TotalDurationMilliseconds));
                SettingsManager.SaveSettings(settings);
            }
        }

        public string Theme
        {
            get => settings.Theme;
            set
            {
                settings.Theme = value;
                OnPropertyChanged(nameof(Theme));
                SettingsManager.SaveSettings(settings);
            }
        }

        public bool IsTopmost
        {
            get => settings.IsTopmost;
            set
            {
                settings.IsTopmost = value;
                OnPropertyChanged(nameof(IsTopmost));
                SettingsManager.SaveSettings(settings);
                if (System.Windows.Application.Current.MainWindow != null)
                {
                    System.Windows.Application.Current.MainWindow.Topmost = value;
                }
            }
        }

        public List<string> RunningApplications
        {
            get => runningApplications;
            private set { runningApplications = value; OnPropertyChanged(nameof(RunningApplications)); }
        }

        public bool IsRunning
        {
            get => isRunning;
            private set { isRunning = value; OnPropertyChanged(nameof(IsRunning)); }
        }

        public string StatusText
        {
            get => statusText;
            set { statusText = value; OnPropertyChanged(nameof(StatusText)); }
        }

        public string StatusColor
        {
            get => statusColor;
            set { statusColor = value; OnPropertyChanged(nameof(StatusColor)); }
        }

        public bool IsMouseMode => ActionType == "Mouse";
        public bool IsKeyboardMode => ActionType == "Keyboard";
        public bool IsClickModeVisible => MouseMode == "Click";
        public bool IsHoldModeVisible => MouseMode == "Hold";
        public bool IsClickDurationMode => MouseMode == "Click" && ClickMode == "Duration";
        public bool IsHoldDurationMode => MouseMode == "Hold" && HoldMode == "HoldDuration";
        public bool IsPressModeVisible => KeyboardMode == "Press";
        public bool IsHoldModeVisibleKeyboard => KeyboardMode == "Hold";
        public bool IsKeyboardHoldDurationMode => KeyboardMode == "Hold" && settings.KeyboardHoldDuration != TimeSpan.Zero;
        public bool IsTimerMode => Mode == "Timer";
        public bool IsRestrictedMode => ClickScope == "Restricted";

        public RelayCommand ToggleAutomationCommand { get; }
        public RelayCommand SetTriggerKeyCommand { get; }
        public RelayCommand SetKeyCommand { get; }
        public RelayCommand ResetSettingsCommand { get; }
        public RelayCommand SetConstantCommand { get; }
        public RelayCommand SetHoldDurationCommand { get; }
        public RelayCommand RefreshApplicationsCommand { get; }
        public RelayCommand OpenOptionsCommand { get; }

        private async Task ManageAutomationAsync(bool start, string stopMessage = null)
        {
            if (start == isRunning)
            {
                Console.WriteLine($"Automation already {(start ? "running" : "stopped")}, no action needed.");
                return;
            }

            if (!start)
            {
                if (automationTcs != null)
                {
                    Console.WriteLine($"Stopping automation: {stopMessage ?? "User requested stop"}");
                    await automationController.StopAutomationAsync(stopMessage ?? "Automation stopped");
                    try
                    {
                        await Task.Delay(100);
                        automationTcs.TrySetResult(true);
                        await automationTcs.Task;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error waiting for automation to stop: {ex.Message}");
                    }
                    automationTcs = null;
                }
                StopTimers();
                UpdateStatus("Not running", "Red");
            }
            else
            {
                if (automationTcs != null)
                {
                    Console.WriteLine("Previous automation still running, waiting for it to stop.");
                    try
                    {
                        await Task.Delay(100);
                        await automationTcs.Task;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error waiting for previous automation to stop: {ex.Message}");
                    }
                    automationTcs = null;
                }

                automationTcs = new TaskCompletionSource<bool>();
                StartTimers();
                UpdateStatus("Running", "Green");
                try
                {
                    await automationController.StartAutomation();
                    automationTcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Automation failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Information);
                    automationTcs.TrySetException(ex);
                }
                finally
                {
                    await ManageAutomationAsync(false, "Automation completed");
                }
            }
        }

        private async Task ToggleAutomation()
        {
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                await ManageAutomationAsync(!isRunning);
            });
        }

        private async Task SetTriggerKey()
        {
            try
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    isSettingToggleKey = true;
                    isSettingKeyboardKey = false;
                    capturedKey = Key.None;
                    var dialog = new Window
                    {
                        Title = "Set Toggle Key",
                        Content = new System.Windows.Controls.TextBlock
                        {
                            Text = "Press a key to set the toggle key. Press Esc to cancel.",
                            Margin = new Thickness(10)
                        },
                        Width = 300,
                        Height = 100,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        Owner = System.Windows.Application.Current.MainWindow
                    };
                    // Apply the current theme
                    dialog.Resources.MergedDictionaries.Add(new ResourceDictionary
                    {
                        Source = new Uri($"/Themes/{settings.Theme}Theme.xaml", UriKind.Relative)
                    });
                    dialog.Show();
                    dialog.KeyDown += (s, e) =>
                    {
                        if (e.Key == Key.Escape)
                        {
                            capturedKey = Key.None;
                            isSettingToggleKey = false;
                            dialog.Close();
                            return;
                        }
                        capturedKey = e.Key;
                        if (e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl))
                            TriggerKeyModifiers |= ModifierKeys.Control;
                        if (e.KeyboardDevice.IsKeyDown(Key.LeftShift) || e.KeyboardDevice.IsKeyDown(Key.RightShift))
                            TriggerKeyModifiers |= ModifierKeys.Shift;
                        if (e.KeyboardDevice.IsKeyDown(Key.LeftAlt) || e.KeyboardDevice.IsKeyDown(Key.RightAlt))
                            TriggerKeyModifiers |= ModifierKeys.Alt;
                        TriggerKey = capturedKey;
                        isSettingToggleKey = false;
                        dialog.Close();
                    };
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in SetTriggerKey: {ex.Message}");
                System.Windows.MessageBox.Show($"Failed to set toggle key: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task SetKey()
        {
            try
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    isSettingKeyboardKey = true;
                    isSettingToggleKey = false;
                    capturedKey = Key.None;
                    var dialog = new Window
                    {
                        Title = "Set Key",
                        Content = new System.Windows.Controls.TextBlock
                        {
                            Text = "Press a key to set the keyboard key. Press Esc to cancel.",
                            Margin = new Thickness(10)
                        },
                        Width = 300,
                        Height = 100,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        Owner = System.Windows.Application.Current.MainWindow
                    };
                    // Apply the current theme
                    dialog.Resources.MergedDictionaries.Add(new ResourceDictionary
                    {
                        Source = new Uri($"/Themes/{settings.Theme}Theme.xaml", UriKind.Relative)
                    });
                    dialog.Show();
                    dialog.KeyDown += (s, e) =>
                    {
                        if (e.Key == Key.Escape)
                        {
                            capturedKey = Key.None;
                            isSettingKeyboardKey = false;
                            dialog.Close();
                            return;
                        }
                        capturedKey = e.Key;
                        KeyboardKey = capturedKey;
                        OnPropertyChanged(nameof(KeyboardKeyDisplay));
                        isSettingKeyboardKey = false;
                        dialog.Close();
                    };
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in SetKey: {ex.Message}");
                System.Windows.MessageBox.Show($"Failed to set key: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task OpenOptionsDialog()
        {
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var dialog = new OptionsDialog(this)
                {
                    Owner = System.Windows.Application.Current.MainWindow
                };
                dialog.ShowDialog();
            });
        }

        private void SetConstant(object _)
        {
            settings.KeyboardHoldDuration = TimeSpan.Zero;
            OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
            SettingsManager.SaveSettings(settings);
        }

        private void SetHoldDuration(object _)
        {
            if (settings.KeyboardHoldDuration == TimeSpan.Zero)
            {
                settings.KeyboardHoldDuration = TimeSpan.FromSeconds(1);
                KeyboardHoldDurationMinutes = settings.KeyboardHoldDuration.Minutes;
                KeyboardHoldDurationSeconds = settings.KeyboardHoldDuration.Seconds;
                KeyboardHoldDurationMilliseconds = settings.KeyboardHoldDuration.Milliseconds;
            }
            OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
            SettingsManager.SaveSettings(settings);
        }

        private void RefreshApplications(object _)
        {
            var apps = applicationDetector.GetRunningApplications();
            Console.WriteLine($"Found {apps.Count} running applications: {string.Join(", ", apps)}");
            RunningApplications = apps;
        }

        private async Task ResetSettings()
        {
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                await ManageAutomationAsync(false, "Resetting settings");

                settings.ClickScope = "Global";
                settings.TargetApplication = "";
                settings.ActionType = "Mouse";
                settings.MouseButton = "Left";
                settings.ClickType = "Single";
                settings.MouseMode = "Click";
                settings.ClickMode = "Constant";
                settings.ClickDuration = TimeSpan.Zero;
                settings.MouseHoldDuration = TimeSpan.FromSeconds(1);
                settings.HoldMode = "ConstantHold";
                settings.MousePhysicalHoldMode = false;
                settings.KeyboardKey = Key.Space;
                settings.KeyboardMode = "Press";
                settings.KeyboardHoldDuration = TimeSpan.Zero;
                settings.KeyboardPhysicalHoldMode = false;
                settings.TriggerKey = Key.F5;
                settings.TriggerKeyModifiers = ModifierKeys.None;
                settings.Interval = TimeSpan.FromSeconds(2);
                settings.Mode = "Constant";
                settings.TotalDuration = TimeSpan.Zero;
                settings.Theme = "Light";
                settings.IsTopmost = false;

                SettingsManager.SaveSettings(settings);

                if (window != null)
                {
                    hotkeyManager?.Dispose();
                    hotkeyManager = new HotkeyManager(window, this);
                    hotkeyManager.RegisterTriggerHotkey(settings.TriggerKey, settings.TriggerKeyModifiers);
                }

                OnPropertyChanged(nameof(ClickScope));
                OnPropertyChanged(nameof(TargetApplication));
                OnPropertyChanged(nameof(ActionType));
                OnPropertyChanged(nameof(MouseButton));
                OnPropertyChanged(nameof(ClickType));
                OnPropertyChanged(nameof(MouseMode));
                OnPropertyChanged(nameof(ClickMode));
                OnPropertyChanged(nameof(ClickDurationMinutes));
                OnPropertyChanged(nameof(ClickDurationSeconds));
                OnPropertyChanged(nameof(ClickDurationMilliseconds));
                OnPropertyChanged(nameof(MouseHoldDurationMinutes));
                OnPropertyChanged(nameof(MouseHoldDurationSeconds));
                OnPropertyChanged(nameof(MouseHoldDurationMilliseconds));
                OnPropertyChanged(nameof(HoldMode));
                OnPropertyChanged(nameof(MousePhysicalHoldMode));
                OnPropertyChanged(nameof(KeyboardKey));
                OnPropertyChanged(nameof(KeyboardKeyDisplay));
                OnPropertyChanged(nameof(KeyboardMode));
                OnPropertyChanged(nameof(KeyboardHoldDurationMinutes));
                OnPropertyChanged(nameof(KeyboardHoldDurationSeconds));
                OnPropertyChanged(nameof(KeyboardHoldDurationMilliseconds));
                OnPropertyChanged(nameof(KeyboardPhysicalHoldMode));
                OnPropertyChanged(nameof(TriggerKey));
                OnPropertyChanged(nameof(TriggerKeyDisplay));
                OnPropertyChanged(nameof(TriggerKeyModifiers));
                OnPropertyChanged(nameof(IntervalMinutes));
                OnPropertyChanged(nameof(IntervalSeconds));
                OnPropertyChanged(nameof(IntervalMilliseconds));
                OnPropertyChanged(nameof(Mode));
                OnPropertyChanged(nameof(TotalDurationMinutes));
                OnPropertyChanged(nameof(TotalDurationSeconds));
                OnPropertyChanged(nameof(TotalDurationMilliseconds));
                OnPropertyChanged(nameof(Theme));
                OnPropertyChanged(nameof(IsTopmost));
                OnPropertyChanged(nameof(IsMouseMode));
                OnPropertyChanged(nameof(IsKeyboardMode));
                OnPropertyChanged(nameof(IsClickModeVisible));
                OnPropertyChanged(nameof(IsHoldModeVisible));
                OnPropertyChanged(nameof(IsClickDurationMode));
                OnPropertyChanged(nameof(IsHoldDurationMode));
                OnPropertyChanged(nameof(IsPressModeVisible));
                OnPropertyChanged(nameof(IsHoldModeVisibleKeyboard));
                OnPropertyChanged(nameof(IsKeyboardHoldDurationMode));
                OnPropertyChanged(nameof(IsTimerMode));
                OnPropertyChanged(nameof(IsRestrictedMode));

                Console.WriteLine("Settings reset to default, HotkeyManager reinitialized.");
            });
        }

        public void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.OriginalSource is System.Windows.Controls.TextBox)
            {
                return;
            }

            if (isSettingToggleKey)
            {
                if (e.Key == Key.Escape)
                {
                    capturedKey = Key.None;
                    isSettingToggleKey = false;
                    return;
                }
                capturedKey = e.Key;
                if (e.KeyboardDevice.IsKeyDown(Key.LeftCtrl) || e.KeyboardDevice.IsKeyDown(Key.RightCtrl))
                    TriggerKeyModifiers |= ModifierKeys.Control;
                if (e.KeyboardDevice.IsKeyDown(Key.LeftShift) || e.KeyboardDevice.IsKeyDown(Key.RightShift))
                    TriggerKeyModifiers |= ModifierKeys.Shift;
                if (e.KeyboardDevice.IsKeyDown(Key.LeftAlt) || e.KeyboardDevice.IsKeyDown(Key.RightAlt))
                    TriggerKeyModifiers |= ModifierKeys.Alt;
                TriggerKey = capturedKey;
                isSettingToggleKey = false;
            }
            else if (isSettingKeyboardKey)
            {
                if (e.Key == Key.Escape)
                {
                    capturedKey = Key.None;
                    isSettingKeyboardKey = false;
                    return;
                }
                capturedKey = e.Key;
                KeyboardKey = capturedKey;
                OnPropertyChanged(nameof(KeyboardKeyDisplay));
                isSettingKeyboardKey = false;
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}