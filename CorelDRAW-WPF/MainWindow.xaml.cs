using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace CorelDRAW_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Controller controller;
        CancellationTokenSource cts;
        public MainWindow()
        {
            InitializeComponent();
            controller = new Controller(this);
        }

        private void ProcessExcelFile_Click(object sender, RoutedEventArgs e)
        {
            ProcessExcelFile.IsEnabled = false;
            ProcessCorelDRAWFile.IsEnabled = false;
            try
            {
                controller.ExtractDataFromExcel();
            }
            finally
            {
                ProcessExcelFile.IsEnabled = true;
                ProcessCorelDRAWFile.IsEnabled = true;
            }
        }

        private void ProcessCorelDRAWFile_Click(object sender, RoutedEventArgs e)
        {
            ProcessExcelFile.IsEnabled = false;
            ProcessCorelDRAWFile.IsEnabled = false;
            try
            {
                controller.InsertDataToCorel();
            }
            finally
            {
                ProcessExcelFile.IsEnabled = true;
                ProcessCorelDRAWFile.IsEnabled = true;
            }
        }

        private async void Test_ClickAsync(object sender, RoutedEventArgs e)
        {
            cts = new CancellationTokenSource();
            OutputText.Text += "Test Async method.\n";
            Task task = controller.GetTaskAsync(cts);
            while(!task.IsCompleted)
            {
                OutputText.Text += "Async method is runing.\n";
                await Task.Delay(1000);
            }
            OutputText.Text += "Async method is completed.\n";
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
