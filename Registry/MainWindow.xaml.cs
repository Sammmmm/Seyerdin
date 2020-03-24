using System;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace Registry.UI {
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e) {
            LibraryChecker.CheckLibraries();
            try {
                await ClientInfo.FetchClientVersion();
            }
            catch (FileLoadException) {

            }
            catch (Exception) {

            }

            var results = await Actions.LoadServersList();

            ListViewServers.DataContext = results;

            await Actions.PingServers(results);
        }

        private void Canvas_MouseDown(object sender, MouseButtonEventArgs e) {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e) {
            e.CanExecute = true;
        }

        private void CloseWindow_Executed(object sender, ExecutedRoutedEventArgs e) {
            SystemCommands.CloseWindow(this);
        }

        private void MinimizeWindow_Executed(object sender, ExecutedRoutedEventArgs e) {
            SystemCommands.MinimizeWindow(this);
        }

        private void ListViewServers_MouseDoubleClick(object sender, MouseButtonEventArgs e) {
            Actions.LaunchServer(ListViewServers.SelectedItem as ServerInfo);
        }

        private void ButtonPlay_Click(object sender, RoutedEventArgs e) {
            var selectedItem = ListViewServers.SelectedItem as ServerInfo ?? ListViewServers.Items[0] as ServerInfo;

            Actions.LaunchServer(selectedItem);
        }
    }
}
