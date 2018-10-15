using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace CorelDRAW_WPF
{
    public partial class MainWindow : Window
    {
        Controller controller;
        CancellationTokenSource cts;
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void ProcessExcelFile_ClickAsync(object sender, RoutedEventArgs e)
        {
            ProcessExcelFile.IsEnabled = false;
            cts = new CancellationTokenSource();
            controller = new Controller(this);
            await controller.StartExcelTaskAsync(cts);
            ProcessExcelFile.IsEnabled = true;
        }

        private async void ProcessCorelDRAWFile_ClickAsync(object sender, RoutedEventArgs e)
        {
            ProcessExcelFile.IsEnabled = false;
            ProcessCorelDRAWFile.IsEnabled = false;
            cts = new CancellationTokenSource();
            await controller.StartCorelTaskAsync(cts);
            ProcessExcelFile.IsEnabled = true;
            ProcessCorelDRAWFile.IsEnabled = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            if (cts != null)
            {
                cts.Cancel();
            }
        }
    }
}
