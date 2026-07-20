// v2.0.3
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;

namespace LeftHandDevice.Views
{
    public sealed partial class SettingsPage : Page
    {
        private AppState? _state;
        private bool _isInitializing;

        public SettingsPage()
        {
            InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);
            _state = e.Parameter as AppState;
            if (_state == null)
                return;

            _isInitializing = true;

            ThemeComboBox.Items.Clear();
            ThemeComboBox.Items.Add(Strings.Get("Settings_Theme_System"));
            ThemeComboBox.Items.Add(Strings.Get("Settings_Theme_Light"));
            ThemeComboBox.Items.Add(Strings.Get("Settings_Theme_Dark"));
            ThemeComboBox.SelectedIndex = (int)_state.Settings.ThemePreference;

            int count = _state.Device.ActiveButtonCount;
            foreach (ComboBoxItem item in ActiveButtonCombo.Items)
            {
                if (item.Tag?.ToString() == count.ToString())
                {
                    ActiveButtonCombo.SelectedItem = item;
                    break;
                }
            }

            WarningSoundToggle.IsOn = _state.Device.WarningSound;
            _isInitializing = false;
        }

        private void ThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitializing || _state == null)
                return;

            if (ThemeComboBox.SelectedIndex < 0)
                return;

            var preference = (AppThemePreference)ThemeComboBox.SelectedIndex;
            if (preference == _state.Settings.ThemePreference)
                return;

            _state.Settings.ThemePreference = preference;
            ThemeService.SetPreference(preference, save: true);
        }

        private void ActiveButtonCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitializing || _state == null)
                return;

            if (ActiveButtonCombo.SelectedItem is ComboBoxItem item
                && int.TryParse(item.Tag?.ToString(), out int val))
            {
                _state.Device.SetActiveButtonCount(val);
            }
        }

        private void WarningSoundToggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (_isInitializing || _state == null)
                return;

            _state.Device.SetWarningSound(WarningSoundToggle.IsOn);
        }
    }
}
