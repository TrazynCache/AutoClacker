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
                        await StopAutomationAsync("Error: Invalid settings. Please check configuration.");
                        return;
                    }

                    Settings settings = viewModel?.CurrentSettings;
                    if (settings == null)
                    {
                        await StopAutomationAsync("Error: Settings not available."); // Should ideally not happen if ValidateSettings passed
                        return;
                    }

                    if (settings.ClickScope == "Restricted" && !ValidateTargetApplication(settings))
                    {
                        string appName = string.IsNullOrEmpty(settings.TargetApplication) ? "the specified application" : settings.TargetApplication;
                        await StopAutomationAsync($"Error: Target application '{appName}' not found or not active.");
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
                else if (settings.HoldMode == "HoldDuration") // Mouse Hold for a specific duration
                {
                    Console.WriteLine($"Holding mouse for {settings.MouseHoldDuration.TotalMilliseconds} ms.");
                    if (settings.MouseAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else // Standard software hold
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
                else if (settings.HoldMode == "ConstantHold") // Mouse Hold indefinitely (toggle)
                {
                    if (!mouseButtonHeld) // This flag is for the standard software constant hold
                    {
                        if (settings.MouseAlternatePhysicalHoldMode)
                        {
                            await SimulateRapidFireMouseHold(settings, token, stopwatch);
                        }
                        else if (settings.MousePhysicalHoldMode)
                        {
                            await SimulatePhysicalMouseHold(settings, token, stopwatch);
                        }
                        else // Standard software constant hold
                        {
                            MouseEventDown(settings);
                            mouseButtonHeld = true;
                        }
                    }
                }
            }
            else // Keyboard Action Type
            {
                if (settings.KeyboardMode == "Press")
                {
                    Console.WriteLine("Performing keyboard press.");
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                    await Task.Delay(50, token);
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                }
                else if (settings.KeyboardHoldDuration != TimeSpan.Zero) // Keyboard Hold for a specific duration
                {
                    Console.WriteLine($"Holding keyboard for {settings.KeyboardHoldDuration.TotalMilliseconds} ms.");
                    if (settings.KeyboardAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else // Standard software hold
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
                else if (settings.KeyboardMode == "Hold" && settings.KeyboardHoldDuration == TimeSpan.Zero) // Keyboard Hold indefinitely (toggle)
                {
                    if (!keyboardKeyHeld) // This flag is for the standard software constant hold
                    {
                        if (settings.KeyboardAlternatePhysicalHoldMode)
                        {
                            await SimulateRapidFireKeyboardHold(settings, token, stopwatch);
                        }
                        else if (settings.KeyboardPhysicalHoldMode)
                        {
                            await SimulatePhysicalKeyboardHold(settings, token, stopwatch);
                        }
                        else // Standard software constant hold
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
                else if (settings.HoldMode == "HoldDuration") // Mouse Hold for a specific duration (Restricted)
                {
                    Console.WriteLine($"Holding mouse (restricted) for {settings.MouseHoldDuration.TotalMilliseconds} ms.");
                    if (settings.MouseAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch, settings.MouseHoldDuration);
                    }
                    else // Standard software hold
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
                else if (settings.HoldMode == "ConstantHold") // Mouse Hold indefinitely (Restricted)
                {
                    // For restricted constant hold, physical/alternate modes are per interval.
                    // Standard software hold would be a single MouseEventDown per interval.
                    if (settings.MouseAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireMouseHold(settings, token, stopwatch);
                    }
                    else if (settings.MousePhysicalHoldMode)
                    {
                        await SimulatePhysicalMouseHold(settings, token, stopwatch);
                    }
                    else // Standard software constant hold (restricted)
                    {
                        MouseEventDown(settings);
                        // Note: Release for this software constant hold in restricted mode is implicitly handled
                        // by the next cycle not re-sending MouseEventDown unless it's a physical type,
                        // or by StopAutomationAsync if automation stops.
                    }
                }
            }
            else // Keyboard Action Type (Restricted)
            {
                if (settings.KeyboardMode == "Press")
                {
                    Console.WriteLine("Performing restricted keyboard press.");
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                    await Task.Delay(50, token);
                    KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 2);
                }
                else if (settings.KeyboardHoldDuration != TimeSpan.Zero) // Keyboard Hold for a specific duration (Restricted)
                {
                    Console.WriteLine($"Holding keyboard (restricted) for {settings.KeyboardHoldDuration.TotalMilliseconds} ms.");
                    if (settings.KeyboardAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch, settings.KeyboardHoldDuration);
                    }
                    else // Standard software hold
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
                else // Keyboard Hold indefinitely (Restricted)
                {
                    if (settings.KeyboardAlternatePhysicalHoldMode)
                    {
                        await SimulateRapidFireKeyboardHold(settings, token, stopwatch);
                    }
                    else if (settings.KeyboardPhysicalHoldMode)
                    {
                        await SimulatePhysicalKeyboardHold(settings, token, stopwatch);
                    }
                    else // Standard software constant hold (restricted)
                    {
                        KeybdEvent((byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey), 0);
                        // Release handled similarly to restricted mouse constant software hold.
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
            TimeSpan eventInterval = TimeSpan.FromMilliseconds(50);

            while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
            {
                MouseEventDown(settings);
                await Task.Delay(eventInterval, cancellationToken); // Original delay was 50ms
            }

            // Original logic: MouseEventUp is called *after* the loop.
            if (!cancellationToken.IsCancellationRequested)
            {
                 // It's possible that the original physical hold didn't even have this explicit up,
                 // relying on the generic StopAutomationAsync. However, for safety, it's better here.
                 // But given the task is to *revert*, I should check original behavior carefully.
                 // The provided code for this revert doesn't show an explicit up here, it was outside.
                 // The previous version before the "correction" had an explicit MouseEventUp(settings)
                 // *after* the loop if not cancelled. Let's stick to that.
                MouseEventUp(settings);
            }
            // The keyboardKeyHeld = false; was indeed an error from the previous change, removing it.
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
            TimeSpan eventInterval = TimeSpan.FromMilliseconds(50);
            byte keyCode = (byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey);

            while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
            {
                KeybdEvent(keyCode, 0); // Key down
                await Task.Delay(eventInterval, cancellationToken); // Original delay was 50ms
            }

            // Original logic: KeybdEvent for key up is called *after* the loop.
            if (!cancellationToken.IsCancellationRequested)
            {
                KeybdEvent(keyCode, NativeMethods.KEYEVENTF_KEYUP); // Key up
            }
        }

        private async Task SimulateRapidFireMouseHold(Settings settings, CancellationToken cancellationToken, Stopwatch stopwatch, TimeSpan? duration = null)
        {
            Console.WriteLine("SimulateRapidFireMouseHold called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in SimulateRapidFireMouseHold. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            TimeSpan pressDuration = TimeSpan.FromMilliseconds(25);
            TimeSpan intervalBetweenPresses = TimeSpan.FromMilliseconds(25);

            try
            {
                while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
                {
                    MouseEventDown(settings);
                    await Task.Delay(pressDuration, cancellationToken);
                    if (cancellationToken.IsCancellationRequested) break;

                    MouseEventUp(settings);
                    if (cancellationToken.IsCancellationRequested) break;

                    await Task.Delay(intervalBetweenPresses, cancellationToken);
                }
            }
            catch (TaskCanceledException) { /* Expected on cancellation */ }
            finally
            {
                // This method simulates rapid clicks, so each click has a down and an up.
                // No specific final MouseEventUp is needed here as the loop completes the action.
                // If cancelled mid-press (after down, before up), StopAutomationAsync handles general cleanup if mouseButtonHeld was used by caller.
                // However, this rapid fire method doesn't use mouseButtonHeld itself.
            }
        }

        private async Task SimulateRapidFireKeyboardHold(Settings settings, CancellationToken cancellationToken, Stopwatch stopwatch, TimeSpan? duration = null)
        {
            Console.WriteLine("SimulateRapidFireKeyboardHold called.");
            if (settings == null)
            {
                Console.WriteLine("settings is null in SimulateRapidFireKeyboardHold. Stopping automation.");
                await StopAutomationAsync("Error: Settings is null");
                return;
            }

            TimeSpan pressDuration = TimeSpan.FromMilliseconds(25);
            TimeSpan intervalBetweenPresses = TimeSpan.FromMilliseconds(25);
            byte keyCode = (byte)KeyInterop.VirtualKeyFromKey(settings.KeyboardKey);

            try
            {
                while (!cancellationToken.IsCancellationRequested && (duration == null || viewModel.GetRemainingDuration() > TimeSpan.Zero))
                {
                    KeybdEvent(keyCode, 0); // Key down
                    await Task.Delay(pressDuration, cancellationToken);
                    if (cancellationToken.IsCancellationRequested) break;

                    KeybdEvent(keyCode, NativeMethods.KEYEVENTF_KEYUP); // Key up
                    if (cancellationToken.IsCancellationRequested) break;

                    await Task.Delay(intervalBetweenPresses, cancellationToken);
                }
            }
            catch (TaskCanceledException) { /* Expected on cancellation */ }
            finally
            {
                // Similar to mouse rapid fire, each action is a full down/up.
                // No specific final key up needed here beyond what the loop does.
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