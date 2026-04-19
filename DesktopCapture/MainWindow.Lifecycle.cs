using System;
using System.ComponentModel;
using System.Windows;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // グローバルホットキーを登録
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            _source = System.Windows.Interop.HwndSource.FromHwnd(helper.Handle);
            _source?.AddHook(HwndHook);

            bool registered = RegisterHotKey(helper.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_C);
            if (!registered)
            {
                MessageBox.Show("ホットキー (Ctrl+Shift+C) の登録に失敗しました。", "警告",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // DPI 変更イベントを購読
            DpiChanged += MainWindow_DpiChanged;

            // 起動時のDPIスケール値を現在値に同期する（物理ピクセル座標の補正は不要）
            if (_isRegionSet)
            {
                var ps = PresentationSource.FromVisual(this);
                double currentDpiX = ps?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                double currentDpiY = ps?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;
                if (Math.Abs(_captureRegionDpiScaleX - currentDpiX) > 0.001 ||
                    Math.Abs(_captureRegionDpiScaleY - currentDpiY) > 0.001)
                {
                    _captureRegionDpiScaleX = currentDpiX;
                    _captureRegionDpiScaleY = currentDpiY;
                    SaveSettings();
                }
            }
        }

        private void MainWindow_Closed(object? sender, EventArgs e)
        {
            // 設定を保存
            SaveMemoToCurrentSavePath();
            SaveSettings();

            // グローバルホットキーを解除
            _source?.RemoveHook(HwndHook);
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            UnregisterHotKey(helper.Handle, HOTKEY_ID);
        }

        private void MainWindow_LocationChanged(object? sender, EventArgs e)
        {
            // 位置変更時は頻繁に保存しないように何もしない
            // 終了時にまとめて保存
        }

        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // サイズ変更時は頻繁に保存しないように何もしない
            // 終了時にまとめて保存
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID)
            {
                Dispatcher.Invoke(() => PerformCapture());
                handled = true;
            }
            return IntPtr.Zero;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // ウィンドウを閉じる前に設定を保存
            SaveMemoToCurrentSavePath();
            SaveSettings();
            base.OnClosing(e);
        }

        private void MainWindow_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            if (!_isRegionSet) return;

            // キャプチャ座標は物理ピクセルで保持しており、DPI変化の影響を受けないため
            // 座標の再計算は不要。DPIスケール値のみ更新する。
            _captureRegionDpiScaleX = e.NewDpi.DpiScaleX;
            _captureRegionDpiScaleY = e.NewDpi.DpiScaleY;
            SaveSettings();
            StatusText.Text = $"DPI変更({e.OldDpi.PixelsPerInchX:F0}→{e.NewDpi.PixelsPerInchX:F0}dpi)。必要に応じてキャプチャ領域を再選択してください。";
        }
    }
}
