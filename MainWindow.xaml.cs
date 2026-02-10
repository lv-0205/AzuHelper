using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using AzuHelper_v2.ViewModels;

namespace AzuHelper_v2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
            Loaded += OnLoaded;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            ApplyDarkTitleBar();
            await _viewModel.EnsureInitializedAsync();
        }

        private void ApplyDarkTitleBar()
        {
            var handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero)
                return;

            const int attributeDarkMode = 20;
            const int attributeDarkModePrev = 19;
            int useDark = 1;

            _ = DwmSetWindowAttribute(handle, attributeDarkMode, ref useDark, sizeof(int));
            _ = DwmSetWindowAttribute(handle, attributeDarkModePrev, ref useDark, sizeof(int));
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int attributeValue, int attributeSize);
    }
}