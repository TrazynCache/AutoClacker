using System;
using System.Threading;
using System.Threading.Tasks;
using AutoClacker.Models;
using AutoClacker.ViewModels;
using AutoClacker.Utilities;
using System.Windows.Input;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AutoClacker.Controllers
{
    public class AutomationController
    {
        private readonly MainViewModel viewModel;
        private CancellationTokenSource cts;
        private readonly ApplicationDetector detector = new ApplicationDetector();
        private bool mouseButtonHeld;
        private bool keyboardKeyHeld;
        private Task currentAutomationTask;
        private readonly object automationLock = new object();
        private volatile bool isRunning;
        private TaskCompletionSource<bool> taskCompletionSource;

        public AutomationController(MainViewModel viewModel)
        {
            this.viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
        }

        public async Task StartAutomation()
        {
            Console.WriteLine($"StartAutomation called. IsRunning: {isRunning}");
            if (isRunning)
            {
                Console.WriteLine("Automation already running. Stopping previous automation.");
                await StopAutomationAsync("Previous automation stopped to start a new one.");
                if (taskCompletionSource != null)
                {
                    try
                    {
                        await Task.Delay(100);
                        await taskCompletionSource.Task;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error waiting for previous automation task to stop: {ex.Message}");
                    }
                }
            }

            lock (automationLock)
            {
                cts = new CancellationTokenSource();
                isRunning = true;
                taskCompletionSource = new TaskCompletionSource<bool>();
            }

            var token = cts.Token;
            currentAutomationTask = Task.Run(async () =>
            {
                try
                {
                    viewModel?.UpdateStatus("Running", "Green");

                    if (!ValidateSettings())
                    {
                        await StopAutomationAsync("Invalid settings detected");
                        return;
                    }

                    Settings settings = viewModel?.CurrentSettings;
                    if (settings == null)
                    {
                        await StopAutomationAsync("Settings is null");
                        return;
                    }

                    if (settings.ClickScope == "Restricted" && !ValidateTargetApplication(settings))
                    {
                        await StopAutomationAsync("Target application not active");
                        return;
                    }

                    TimeSpan effectiveInterval = settings.Interval.TotalMilliseconds < 100 ? TimeSpan.FromMilliseconds(100) : settings.Interval;

                    mouseButtonHeld = false;
                    keyboardKeyHeld = false;

                    var stopwatch = Stopwatch.StartNew();
                    bool shouldBreak = false;

                    while (!token.IsCancellationRequested && !shouldBreak)
                    {
                        var cycleStartTime = stopwatch.Elapsed;

                        if (settings.ClickScope == "Restricted")
                        {
                            await PerformRestrictedAction(settings, stopwatch, token);
                        }
                        else
                        {
                            await PerformGlobalAction(settings, stopwatch, token);
                        }

                        // Check duration-based modes only
                        if ((settings.ActionType == "Mouse" && settings.MouseMode == "Click" && settings.ClickMode == "Duration") ||
                            (settings.ActionType == "Mouse" && settings.MouseMode == "Hold" && settings.HoldMode == "HoldDuration") ||
                            (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Press" && settings.Mode == "Timer") ||
                            (settings.ActionType == "Keyboard" && settings.KeyboardMode == "Hold" && settings.KeyboardHoldDuration != TimeSpan.Zero))
                        {
                            if (viewModel.GetRemainingDuration() <= TimeSpan.Zero)
                            {
                                shouldBreak = true;
                            }
                        }

                        var cycleElapsedTime = stopwatch.Elapsed - cycleStartTime;
                        var remainingCycleTime = effectiveInterval - cycleElapsedTime;
                        if (remainingCycleTime.TotalMilliseconds > 0)
                        {
                            await Task.Delay(remainingCycleTime, token);
                        }
                    }

                    if (mouseButtonHeld)
                    {
                        MouseEventUp(settings);
                        mouseButtonHeld = false;
                        keyboardKeyHeld = false;
                    }
                    if (keyboardKeyHeld)
                    {
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                        keyboardKeyHeld = false;
                    }

                    await StopAutomationAsync("Automation completed");
                }
                catch (TaskCanceledException)
                {
                    await StopAutomationAsync("Automation stopped");
                }
                catch (Exception ex)
                {
                    await StopAutomationAsync($"Error: {ex.Message}");
                }
                finally
                {
                    lock (automationLock)
                    {
                        isRunning = false;
                        taskCompletionSource?.SetResult(true);
                    }
                }
            }, token);

            await currentAutomationTask;
        }

        public async Task StopAutomationAsync(string message = "Not running")
        {
            Console.WriteLine($"StopAutomationAsync called with message: {message}");
            lock (automationLock)
            {
                if (cts != null)
                {
                    cts.Cancel();
                    cts.Dispose();
                    cts = null;
                }

                if (mouseButtonHeld)
                {
                    MouseEventUp(viewModel?.CurrentSettings ?? new Settings { MouseButton = "Left" });
                    mouseButtonHeld = false;
                    keyboardKeyHeld = false;
                }
                if (keyboardKeyHeld && viewModel?.CurrentSettings?.ActionType == "Keyboard")
                {
                    var key = viewModel.CurrentSettings?.KeyboardKey ?? Key.Space;
                    if (key != Key.None)
                    {
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(key), 2);
                    }
                    else
                    {
                        Console.WriteLine("KeyboardKey is Key.None, skipping key release.");
                    }
                    keyboardKeyHeld = false;
                }

                viewModel?.StopTimers();
                viewModel?.UpdateStatus(message, "Red");
                isRunning = false;
                taskCompletionSource?.SetResult(true);
                taskCompletionSource = null;
            }
            await Task.Delay(100);
        }

        private bool ValidateSettings()
        {
            var settings = viewModel?.CurrentSettings;
            if (settings == null)
            {
                Console.WriteLine("Settings object is null.");
                return false;
            }

            if (string.IsNullOrEmpty(settings.ActionType))
            {
                Console.WriteLine("ActionType is null or empty. Setting to default 'Mouse'.");
                settings.ActionType = "Mouse";
            }
            if (string.IsNullOrEmpty(settings.MouseButton))
            {
                Console.WriteLine("MouseButton is null or empty. Setting to default 'Left'.");
                settings.MouseButton = "Left";
            }
            if (string.IsNullOrEmpty(settings.ClickType))
            {
                Console.WriteLine("ClickType is null or empty. Setting to default 'Single'.");
                settings.ClickType = "Single";
            }
            if (string.IsNullOrEmpty(settings.MouseMode))
            {
                Console.WriteLine("MouseMode is null or empty. Setting to default 'Click'.");
                settings.MouseMode = "Click";
            }
            if (string.IsNullOrEmpty(settings.ClickMode))
            {
                Console.WriteLine("ClickMode is null or empty. Setting to default 'Constant'.");
                settings.ClickMode = "Constant";
            }
            if (string.IsNullOrEmpty(settings.HoldMode))
            {
                Console.WriteLine("HoldMode is null or empty. Setting to default 'ConstantHold'.");
                settings.HoldMode = "ConstantHold";
            }
            if (string.IsNullOrEmpty(settings.KeyboardMode))
            {
                Console.WriteLine("KeyboardMode is null or empty. Setting to default 'Press'.");
                settings.KeyboardMode = "Press";
            }
            if (string.IsNullOrEmpty(settings.Mode))
            {
                Console.WriteLine("Mode is null or empty. Setting to default 'Constant'.");
                settings.Mode = "Constant";
            }
            if (settings.MouseHoldDuration == TimeSpan.Zero)
            {
                Console.WriteLine("MouseHoldDuration is zero. Setting to default '1 second'.");
                settings.MouseHoldDuration = TimeSpan.FromSeconds(1);
            }
            if (settings.KeyboardKey == Key.None && settings.ActionType == "Keyboard")
            {
                Console.WriteLine("KeyboardKey is Key.None for keyboard action. Setting to default 'Space'.");
                settings.KeyboardKey = Key.Space;
            }

            return true;
        }

        private bool ValidateTargetApplication(Settings settings)
        {
            var process = detector.GetProcessByName(settings?.TargetApplication);
            bool isValid = process != null && !NativeMethods.IsIconic(process.MainWindowHandle);
            Console.WriteLine($"ValidateTargetApplication: Target={settings?.TargetApplication}, IsValid={isValid}");
            return isValid;
        }

        private async Task PerformGlobalAction(Settings settings, Stopwatch stopwatch, CancellationToken token)
        {
            Console.WriteLine("PerformGlobalAction called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in PerformGlobalAction. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            if (settings.ActionType == "Mouse")
            {
                keyboardKeyHeld = false;
                if (settings.MouseMode == "Click")
                {
                    Console.WriteLine("Performing mouse click.");
                    MouseEventDown(settings);
                    await Task.Delay(10, token);
                    MouseEventUp(settings);
                    if (settings.ClickType == "Double")
                    {
                        Console.WriteLine("Performing double click.");
                        await Task.Delay(50, token);
                        MouseEventDown(settings);
                        await Task.Delay(10, token);
                        MouseEventUp(settings);
                    }
                }
                else if (settings.HoldMode == "HoldDuration")
                {
                    Console.WriteLine($"Holding mouse for {settings.MouseHoldDuration.TotalMilliseconds} ms.");
                    if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else
                    {
                        TimeSpan holdDuration = settings.MouseHoldDuration.TotalMilliseconds < 500 ? TimeSpan.FromMilliseconds(500) : settings.MouseHoldDuration;
                        MouseEventDown(settings);
                        await Task.Delay(holdDuration, token);
                        if (!token.IsCancellationRequested)
                        {
                            MouseEventUp(settings);
                        }
                    }
                }
                else if (settings.HoldMode == "ConstantHold")
                {
                    if (!mouseButtonHeld)
                    {
                        if (settings.MousePhysicalHoldMode)
                        {
                            await SimulatePhysicalMouseHold(settings, token, stopwatch);
                        }
                        else
                        {
                            MouseEventDown(settings);
                            mouseButtonHeld = true;
                        }
                    }
                }
            }
            else
            {
                if (settings.KeyboardMode == "Press")
                {
                    Console.WriteLine("Performing keyboard press.");
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                    await Task.Delay(50, token);
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                }
                else if (settings.KeyboardHoldDuration != TimeSpan.Zero)
                {
                    Console.WriteLine($"Holding keyboard for {settings.KeyboardHoldDuration.TotalMilliseconds} ms.");
                    if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else
                    {
                        TimeSpan holdDuration = settings.KeyboardHoldDuration.TotalMilliseconds < 500 ? TimeSpan.FromMilliseconds(500) : settings.KeyboardHoldDuration;
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                        await Task.Delay(holdDuration, token);
                        if (!token.IsCancellationRequested)
                        {
                            KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                        }
                    }
                }
                else if (settings.KeyboardMode == "Hold" && settings.KeyboardHoldDuration == TimeSpan.Zero)
                {
                    if (!keyboardKeyHeld)
                    {
                        if (settings.KeyboardPhysicalHoldMode)
                        {
                            await SimulatePhysicalKeyboardHold(settings, token, stopwatch);
                        }
                        else
                        {
                            KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                            keyboardKeyHeld = true;
                        }
                    }
                }
            }
        }

        private async Task PerformRestrictedAction(Settings settings, Stopwatch stopwatch, CancellationToken token)
        {
            Console.WriteLine("PerformRestrictedAction called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in PerformRestrictedAction. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            var process = detector.GetProcessByName(settings.TargetApplication);
            if (process == null || NativeMethods.IsIconic(process.MainWindowHandle))
            {
                await StopAutomationAsync("Target application not active");
                return;
            }

            // Activate target application window
            NativeMethods.SetForegroundWindow(process.MainWindowHandle);
            await Task.Delay(50, token); // Allow time for window activation

            NativeMethods.GetClientRect(process.MainWindowHandle, out NativeMethods.RECT rect);
            int x = (rect.Right - rect.Left) / 2;
            int y = (rect.Bottom - rect.Top) / 2;

            if (settings.ActionType == "Mouse")
            {
                keyboardKeyHeld = false;
                if (settings.MouseMode == "Click")
                {
                    Console.WriteLine("Performing restricted mouse click.");
                    MouseEventDown(settings);
                    await Task.Delay(10, token);
                    MouseEventUp(settings);
                    if (settings.ClickType == "Double")
                    {
                        Console.WriteLine("Performing restricted double click.");
                        await Task.Delay(50, token);
                        MouseEventDown(settings);
                        await Task.Delay(10, token);
                        MouseEventUp(settings);
                    }
                }
                else if (settings.HoldMode == "HoldDuration")
                {
                    Console.WriteLine($"Holding mouse (restricted) for {settings.MouseHoldDuration.TotalMilliseconds} ms.");
                    if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else
                    {
                        TimeSpan holdDuration = settings.MouseHoldDuration.TotalMilliseconds < 500 ? TimeSpan.FromMilliseconds(500) : settings.MouseHoldDuration;
                        MouseEventDown(settings);
                        await Task.Delay(holdDuration, token);
                        if (!token.IsCancellationRequested)
                        {
                            MouseEventUp(settings);
                        }
                    }
                }
                else if (settings.HoldMode == "ConstantHold")
                {
                    if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch);
                    }
                    else
                    {
                        MouseEventDown(settings);
                    }
                }
            }
            else
            {
                if (settings.KeyboardMode == "Press")
                {
                    Console.WriteLine("Performing restricted keyboard press.");
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                    await Task.Delay(50, token);
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                }
                else if (settings.KeyboardHoldDuration != TimeSpan.Zero)
                {
                    Console.WriteLine($"Holding keyboard (restricted) for {settings.KeyboardHoldDuration.TotalMilliseconds} ms.");
                    if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else
                    {
                        TimeSpan holdDuration = settings.KeyboardHoldDuration.TotalMilliseconds < 500 ? TimeSpan.FromMilliseconds(500) : settings.KeyboardHoldDuration;
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                        await Task.Delay(holdDuration, token);
                        if (!token.IsCancellationRequested)
                        {
                            KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                        }
                    }
                }
                else
                {
                    if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch);
                    }
                    else
                    {
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                    }
                }
            }
        }

        private async Task SimulatePhysicalMouseHold(Settings settings, CancellationToken cancellationToken, Stopwatch stopwatch, TimeSpan? duration = null)
        {
            Console.WriteLine("SimulatePhysicalMouseHold called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in SimulatePhysicalMouseHold. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            TimeSpan holdDuration = duration ?? TimeSpan.MaxValue;
            TimeSpan eventInterval = TimeSpan.FromMilliseconds(50); // Reduced frequency

            while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
            {
                MouseEventDown(settings);
                await Task.Delay(50, cancellationToken);
            }

            if (!cancellationToken.IsCancellationRequested)
            {
                MouseEventUp(settings);
            }
            keyboardKeyHeld = false;
        }

        private async Task SimulatePhysicalKeyboardHold(Settings settings, CancellationToken cancellationToken, Stopwatch stopwatch, TimeSpan? duration = null)
        {
            Console.WriteLine("SimulatePhysicalKeyboardHold called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in SimulatePhysicalKeyboardHold. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            TimeSpan holdDuration = duration ?? TimeSpan.MaxValue;
            TimeSpan eventInterval = TimeSpan.FromMilliseconds(50); // Reduced frequency

            while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
            {
                KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                await Task.Delay(50, cancellationToken);
            }

            if (!cancellationToken.IsCancellationRequested)
            {
                KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
            }
        }

        private void MouseEventDown(Settings settings)
        {
            if (settings == null)
            {
                Console.WriteLine("settings is null in MouseEventDown. Cannot proceed.");
                return;
            }

            if (string.IsNullOrEmpty(settings.MouseButton))
            {
                Console.WriteLine("MouseButton is null or empty in MouseEventDown. Using default 'Left'.");
                settings.MouseButton = "Left";
            }

            Console.WriteLine($"MouseEventDown called for {settings.MouseButton}.");
            
            uint downFlag;
            switch (settings.MouseButton)
            {
                case "Left":
                    downFlag = NativeMethods.MOUSEEVENTF_LEFTDOWN;
                    break;
                case "Right":
                    downFlag = NativeMethods.MOUSEEVENTF_RIGHTDOWN;
                    break;
                case "Middle":
                    downFlag = NativeMethods.MOUSEEVENTF_MIDDLEDOWN;
                    break;
                default:
                    Console.WriteLine($"Invalid MouseButton value '{settings.MouseButton}' in MouseEventDown. Using default 'Left'.");
                    downFlag = NativeMethods.MOUSEEVENTF_LEFTDOWN;
                    break;
            }

            var input = new NativeMethods.INPUT
            {
                Type = NativeMethods.INPUT_MOUSE,
                Data = new NativeMethods.InputUnion
                {
                    Mouse = new NativeMethods.MOUSEINPUT
                    {
                        dwFlags = downFlag,
                        dx = 0,
                        dy = 0,
                        MouseData = 0,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            NativeMethods.SendInput(1, new[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
        }

        private void MouseEventUp(Settings settings)
        {
            if (settings == null)
            {
                Console.WriteLine("settings is null in MouseEventUp. Cannot proceed.");
                return;
            }

            if (string.IsNullOrEmpty(settings.MouseButton))
            {
                Console.WriteLine("MouseButton is null or empty in MouseEventUp. Using default 'Left'.");
                settings.MouseButton = "Left";
            }

            Console.WriteLine($"MouseEventUp called for {settings.MouseButton}.");
            
            uint upFlag;
            switch (settings.MouseButton)
            {
                case "Left":
                    upFlag = NativeMethods.MOUSEEVENTF_LEFTUP;
                    break;
                case "Right":
                    upFlag = NativeMethods.MOUSEEVENTF_RIGHTUP;
                    break;
                case "Middle":
                    upFlag = NativeMethods.MOUSEEVENTF_MIDDLEUP;
                    break;
                default:
                    Console.WriteLine($"Invalid MouseButton value '{settings.MouseButton}' in MouseEventUp. Using default 'Left'.");
                    upFlag = NativeMethods.MOUSEEVENTF_LEFTUP;
                    break;
            }

            var input = new NativeMethods.INPUT
            {
                Type = NativeMethods.INPUT_MOUSE,
                Data = new NativeMethods.InputUnion
                {
                    Mouse = new NativeMethods.MOUSEINPUT
                    {
                        dwFlags = upFlag,
                        dx = 0,
                        dy = 0,
                        MouseData = 0,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            NativeMethods.SendInput(1, new[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
        }

        private void KeybdEvent(byte key, uint flags)
        {
            Console.WriteLine($"KeybdEvent called: Key={key}, Flags={flags}.");
            var input = new NativeMethods.INPUT
            {
                Type = NativeMethods.INPUT_KEYBOARD,
                Data = new NativeMethods.InputUnion
                {
                    Keyboard = new NativeMethods.KEYBDINPUT
                    {
                        wVk = key,
                        wScan = 0,
                        dwFlags = flags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            NativeMethods.SendInput(1, new[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
        }
    }
}