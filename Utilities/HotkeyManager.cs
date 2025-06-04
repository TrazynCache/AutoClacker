using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using AutoClacker.ViewModels;

namespace AutoClacker.Utilities
{
    public class HotkeyManager
    {
        private const int WM_HOTKEY = 0x0312;
        private const int MOD_ALT = 0x0001;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;
        private const int MOD_WIN = 0x0008;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private readonly Window window;
        private readonly MainViewModel viewModel;
        private IntPtr hWnd;
        private int hotkeyId = 1;

        public HotkeyManager(Window window, MainViewModel viewModel)
        {
            this.window = window;
            this.viewModel = viewModel;
            var source = PresentationSource.FromVisual(window) as HwndSource;
            if (source == null)
            {
                throw new InvalidOperationException("Failed to get HwndSource from window.");
            }
            hWnd = source.Handle;
            source.AddHook(WndProc);
            Console.WriteLine("HotkeyManager initialized successfully.");
        }

        public void RegisterTriggerHotkey(Key key, ModifierKeys modifiers)
        {
            UnregisterHotKey(hWnd, hotkeyId);
            uint fsModifiers = 0;
            if (modifiers.HasFlag(ModifierKeys.Alt)) fsModifiers |= MOD_ALT;
            if (modifiers.HasFlag(ModifierKeys.Control)) fsModifiers |= MOD_CONTROL;
            if (modifiers.HasFlag(ModifierKeys.Shift)) fsModifiers |= MOD_SHIFT;
            if (modifiers.HasFlag(ModifierKeys.Windows)) fsModifiers |= MOD_WIN;

            uint vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            bool success = RegisterHotKey(hWnd, hotkeyId, fsModifiers, vk);
            Console.WriteLine($"Hotkey registration: Key={key}, Modifiers={fsModifiers}, Success={success}");
            if (!success)
            {
                // Log the error with more details if possible (e.g., Marshal.GetLastWin32Error())
                Console.WriteLine($"Failed to register hotkey: Key={key}, Modifiers={fsModifiers}. Error Code: {Marshal.GetLastWin32Error()}. It may be in use by another application.");
            }
            return success;
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY && wParam.ToInt32() == hotkeyId)
            {
                Console.WriteLine("Hotkey triggered, executing ToggleAutomationCommand.");
                viewModel.ToggleAutomationCommand.Execute(null);
                handled = true;
            }
            return IntPtr.Zero;
        }

        public void Dispose()
        {
            UnregisterHotKey(hWnd, hotkeyId);
            Console.WriteLine("HotkeyManager disposed, unregistered hotkey.");
        }
    }
}