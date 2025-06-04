using System;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Windows.Input;

namespace AutoClacker.Models
{
    [DataContract]
    public class Settings : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string clickScope = "Global";
        private string targetApplication = "";
        private string actionType = "Mouse";
        private string mouseButton = "Left";
        private string clickType = "Single";
        private string mouseMode = "Click";
        private string clickMode = "Constant";
        private TimeSpan clickDuration = TimeSpan.Zero;
        private TimeSpan mouseHoldDuration = TimeSpan.FromSeconds(1);
        private string holdMode = "HoldDuration";
        private bool mousePhysicalHoldMode = false;
        private int keyboardKey = 0;
        private string keyboardMode = "Press";
        private TimeSpan keyboardHoldDuration = TimeSpan.Zero;
        private bool keyboardPhysicalHoldMode = false;
        private int triggerKey = 116;
        private int triggerKeyModifiers = 0;
        private TimeSpan interval = TimeSpan.FromSeconds(2);
        private string mode = "Constant";
        private TimeSpan totalDuration = TimeSpan.Zero;
        private string theme = "Light";
        private bool isTopmost = false;
        private bool mouseAlternatePhysicalHoldMode = false;
        private bool keyboardAlternatePhysicalHoldMode = false;

        [DataMember]
        public string ClickScope
        {
            get => clickScope;
            set { clickScope = value; OnPropertyChanged(nameof(ClickScope)); }
        }

        [DataMember]
        public string TargetApplication
        {
            get => targetApplication;
            set { targetApplication = value; OnPropertyChanged(nameof(TargetApplication)); }
        }

        [DataMember]
        public string ActionType
        {
            get => actionType;
            set { actionType = value; OnPropertyChanged(nameof(ActionType)); }
        }

        [DataMember]
        public string MouseButton
        {
            get => mouseButton;
            set { mouseButton = value; OnPropertyChanged(nameof(MouseButton)); }
        }

        [DataMember]
        public string ClickType
        {
            get => clickType;
            set { clickType = value; OnPropertyChanged(nameof(ClickType)); }
        }

        [DataMember]
        public string MouseMode
        {
            get => mouseMode;
            set { mouseMode = value; OnPropertyChanged(nameof(MouseMode)); }
        }

        [DataMember]
        public string ClickMode
        {
            get => clickMode;
            set { clickMode = value; OnPropertyChanged(nameof(ClickMode)); }
        }

        [DataMember]
        public TimeSpan ClickDuration
        {
            get => clickDuration;
            set { clickDuration = value; OnPropertyChanged(nameof(ClickDuration)); }
        }

        [DataMember]
        public TimeSpan MouseHoldDuration
        {
            get => mouseHoldDuration;
            set { mouseHoldDuration = value; OnPropertyChanged(nameof(MouseHoldDuration)); }
        }

        [DataMember]
        public string HoldMode
        {
            get => holdMode;
            set { holdMode = value; OnPropertyChanged(nameof(HoldMode)); }
        }

        [DataMember]
        public bool MousePhysicalHoldMode
        {
            get => mousePhysicalHoldMode;
            set { mousePhysicalHoldMode = value; OnPropertyChanged(nameof(MousePhysicalHoldMode)); }
        }

        [DataMember]
        public Key KeyboardKey
        {
            get => (Key)keyboardKey;
            set { keyboardKey = (int)value; OnPropertyChanged(nameof(KeyboardKey)); }
        }

        [DataMember]
        public string KeyboardMode
        {
            get => keyboardMode;
            set { keyboardMode = value; OnPropertyChanged(nameof(KeyboardMode)); }
        }

        [DataMember]
        public TimeSpan KeyboardHoldDuration
        {
            get => keyboardHoldDuration;
            set { keyboardHoldDuration = value; OnPropertyChanged(nameof(KeyboardHoldDuration)); }
        }

        [DataMember]
        public bool KeyboardPhysicalHoldMode
        {
            get => keyboardPhysicalHoldMode;
            set { keyboardPhysicalHoldMode = value; OnPropertyChanged(nameof(KeyboardPhysicalHoldMode)); }
        }

        [DataMember]
        public Key TriggerKey
        {
            get => (Key)triggerKey;
            set { triggerKey = (int)value; OnPropertyChanged(nameof(TriggerKey)); }
        }

        [DataMember]
        public ModifierKeys TriggerKeyModifiers
        {
            get => (ModifierKeys)triggerKeyModifiers;
            set { triggerKeyModifiers = (int)value; OnPropertyChanged(nameof(TriggerKeyModifiers)); }
        }

        [DataMember]
        public TimeSpan Interval
        {
            get => interval;
            set { interval = value; OnPropertyChanged(nameof(Interval)); }
        }

        [DataMember]
        public string Mode
        {
            get => mode;
            set { mode = value; OnPropertyChanged(nameof(Mode)); }
        }

        [DataMember]
        public TimeSpan TotalDuration
        {
            get => totalDuration;
            set { totalDuration = value; OnPropertyChanged(nameof(TotalDuration)); }
        }

        [DataMember]
        public string Theme
        {
            get => theme;
            set { theme = value; OnPropertyChanged(nameof(Theme)); }
        }

        [DataMember]
        public bool IsTopmost
        {
            get => isTopmost;
            set { isTopmost = value; OnPropertyChanged(nameof(IsTopmost)); }
        }

        [DataMember]
        public bool MouseAlternatePhysicalHoldMode
        {
            get => mouseAlternatePhysicalHoldMode;
            set { mouseAlternatePhysicalHoldMode = value; OnPropertyChanged(nameof(MouseAlternatePhysicalHoldMode)); }
        }

        [DataMember]
        public bool KeyboardAlternatePhysicalHoldMode
        {
            get => keyboardAlternatePhysicalHoldMode;
            set { keyboardAlternatePhysicalHoldMode = value; OnPropertyChanged(nameof(KeyboardAlternatePhysicalHoldMode)); }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}