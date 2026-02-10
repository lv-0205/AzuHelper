using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using AzuHelper_v2.Models;

namespace AzuHelper_v2;

public partial class SettingsDialog : Window
{
    public SettingsDialog(AppConfig config)
    {
        InitializeComponent();
        DataContext = config;
        Loaded += OnLoaded;
    }

    public AppConfig Config => (AppConfig)DataContext;

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        ApplyDarkTitleBar();
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
