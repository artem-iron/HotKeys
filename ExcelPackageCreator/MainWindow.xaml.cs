using System.ComponentModel;
using System.Windows;

namespace ExcelPackageCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MouseKeyHook _keyHook = new MouseKeyHook();

        public MainWindow()
        {
            InitializeComponent();

            _keyHook.Subscribe();

            Closing += OnWindowClosing;

            ReadConfig();
        }

        public void OnWindowClosing(object sender, CancelEventArgs e)
        {
            _keyHook.Unsubscribe();
        }

        private void ReadConfig()
        {
            monitoredPathTextBox.Text = ZipFileCreator.MonitoredPath;
            savePathTextBox.Text = ZipFileCreator.SavePath;
        }

        private void monitoredPathSaveButton_Click(object sender, RoutedEventArgs e)
        {
            ZipFileCreator.StoreNewMonitoredPath(monitoredPathTextBox.Text);

            ReadConfig();
        }

        private void savePathSaveButton_Click(object sender, RoutedEventArgs e)
        {
            ZipFileCreator.StoreNewSavePath(savePathTextBox.Text);

            ReadConfig();
        }
    }
}
