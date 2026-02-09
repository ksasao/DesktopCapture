using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace DesktopCapture
{
    public partial class TemplateHelpWindow : Window
    {
        public string? SelectedTemplate { get; private set; }

        public TemplateHelpWindow()
        {
            InitializeComponent();
        }

        private void TemplateExample_MouseEnter(object sender, MouseEventArgs e)
        {
            if (sender is Border border)
            {
                border.Background = new SolidColorBrush(Color.FromRgb(230, 240, 255));
                border.BorderBrush = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            }
        }

        private void TemplateExample_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sender is Border border)
            {
                border.Background = Brushes.White;
                border.BorderBrush = Brushes.LightGray;
            }
        }

        private void TemplateExample_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Tag is string template)
            {
                try
                {
                    Clipboard.SetText(template);
                    SelectedTemplate = template;
                    StatusTextBlock.Text = $"コピーしました: {template}";
                    
                    // 3秒後にステータスをクリア
                    var timer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(3)
                    };
                    timer.Tick += (s, args) =>
                    {
                        StatusTextBlock.Text = "";
                        timer.Stop();
                    };
                    timer.Start();
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = $"コピーに失敗しました: {ex.Message}";
                    StatusTextBlock.Foreground = Brushes.Red;
                }
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}