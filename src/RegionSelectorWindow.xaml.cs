using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace DesktopCapture
{
    public partial class RegionSelectorWindow : Window
    {
        private System.Windows.Point _startPoint;
        private bool _isSelecting = false;
        public Rectangle SelectedRegion { get; private set; }

        public RegionSelectorWindow()
        {
            InitializeComponent();
            InfoText.Text = "マウスをドラッグして領域を選択してください (ESCキーでキャンセル)";
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                _startPoint = e.GetPosition(this);
                _isSelecting = true;
                SelectionRectangle.Visibility = Visibility.Visible;
            }
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isSelecting)
            {
                var currentPoint = e.GetPosition(this);
                
                double x = Math.Min(_startPoint.X, currentPoint.X);
                double y = Math.Min(_startPoint.Y, currentPoint.Y);
                double width = Math.Abs(currentPoint.X - _startPoint.X);
                double height = Math.Abs(currentPoint.Y - _startPoint.Y);

                Canvas.SetLeft(SelectionRectangle, x);
                Canvas.SetTop(SelectionRectangle, y);
                SelectionRectangle.Width = width;
                SelectionRectangle.Height = height;

                InfoText.Text = $"領域: {width:F0} x {height:F0} (マウスを離して確定)";
            }
        }

        private void Window_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSelecting)
            {
                _isSelecting = false;
                
                var currentPoint = e.GetPosition(this);
                
                double wpfX = Math.Min(_startPoint.X, currentPoint.X);
                double wpfY = Math.Min(_startPoint.Y, currentPoint.Y);
                double wpfWidth = Math.Abs(currentPoint.X - _startPoint.X);
                double wpfHeight = Math.Abs(currentPoint.Y - _startPoint.Y);

                if (wpfWidth > 10 && wpfHeight > 10)
                {
                    // WPF座標をスクリーン座標に変換
                    var screenPoint = PointToScreen(new System.Windows.Point(wpfX, wpfY));
                    
                    // DPIスケーリングを取得
                    var source = PresentationSource.FromVisual(this);
                    double dpiX = 1.0;
                    double dpiY = 1.0;
                    if (source?.CompositionTarget != null)
                    {
                        dpiX = source.CompositionTarget.TransformToDevice.M11;
                        dpiY = source.CompositionTarget.TransformToDevice.M22;
                    }

                    // 物理ピクセルに変換
                    int x = (int)screenPoint.X;
                    int y = (int)screenPoint.Y;
                    int width = (int)(wpfWidth * dpiX);
                    int height = (int)(wpfHeight * dpiY);

                    SelectedRegion = new Rectangle(x, y, width, height);
                    DialogResult = true;
                    Close();
                }
                else
                {
                    MessageBox.Show("領域が小さすぎます。もう一度選択してください。", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    SelectionRectangle.Visibility = Visibility.Collapsed;
                    InfoText.Text = "マウスをドラッグして領域を選択してください (ESCキーでキャンセル)";
                }
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}