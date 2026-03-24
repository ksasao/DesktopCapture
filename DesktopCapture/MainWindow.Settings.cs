using System;
using System.IO;
using System.Windows;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private void LoadSettings()
        {
            _isLoadingSettings = true;
            try
            {
                // 保存先パスを復元（空の場合はデフォルト）
                if (!string.IsNullOrEmpty(_settings.SavePath) && Directory.Exists(_settings.SavePath))
                {
                    SavePathTextBox.Text = _settings.SavePath;
                }
                else
                {
                    SavePathTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                }
                _lastAppliedSavePath = NormalizeSavePath(SavePathTextBox.Text);

                // ファイル名テンプレートを決定（空の場合はデフォルト値）
                string currentTemplate = string.IsNullOrWhiteSpace(_settings.FileNameTemplate)
                    ? "cap_{yyyyMMdd_HHmmss}_{###}"
                    : _settings.FileNameTemplate;

                // ファイル名履歴を復元（現在のテンプレートが含まれていない場合は先頭に追加）
                FileNameTemplateComboBox.Items.Clear();

                // 現在のテンプレートが履歴に含まれているか確認
                bool currentTemplateInHistory = _settings.FileNameHistory.Contains(currentTemplate);

                // 現在のテンプレートを先頭に追加（履歴になければ）
                if (!currentTemplateInHistory && !string.IsNullOrWhiteSpace(currentTemplate))
                {
                    FileNameTemplateComboBox.Items.Add(currentTemplate);
                }

                // 履歴を追加
                foreach (var history in _settings.FileNameHistory)
                {
                    FileNameTemplateComboBox.Items.Add(history);
                }

                // 現在のテンプレートを設定
                FileNameTemplateComboBox.Text = currentTemplate;

                // 画像形式を復元（この時点でファイル名が設定されているので上書きされない）
                FormatComboBox.SelectedIndex = _settings.ImageFormat;

                // クリップボードコピー設定を復元
                CopyToClipboardCheckBox.IsChecked = _settings.CopyToClipboard;

                // キャプチャ領域を復元
                if (_settings.CaptureRegion != null)
                {
                    _captureRegion = _settings.CaptureRegion.ToRectangle();
                    _captureRegionDpiScaleX = _settings.CaptureRegion.DpiScaleX > 0 ? _settings.CaptureRegion.DpiScaleX : 1.0;
                    _captureRegionDpiScaleY = _settings.CaptureRegion.DpiScaleY > 0 ? _settings.CaptureRegion.DpiScaleY : 1.0;
                    _isRegionSet = true;

                    RegionText.Text = $"X:{_captureRegion.X}, Y:{_captureRegion.Y}, " +
                                      $"W:{_captureRegion.Width}, H:{_captureRegion.Height}";
                    CaptureButton.IsEnabled = true;
                    StatusText.Text = "前回の設定を復元しました。キャプチャボタンまたはCtrl+Shift+Cでキャプチャできます。";
                }
            }
            finally
            {
                _isLoadingSettings = false;
            }

            LoadMemoAndHistoryFromCurrentSavePath();
            _captureCount = ScanMaxCaptureCount(SavePathTextBox.Text);
            _lastAppliedMemoPath = GetMemoFilePath();
        }

        private void LoadWindowPosition()
        {
            // ウィンドウ位置とサイズを復元
            if (_settings.WindowWidth.HasValue && _settings.WindowHeight.HasValue)
            {
                Width = _settings.WindowWidth.Value;
                Height = _settings.WindowHeight.Value;
            }

            if (_settings.WindowLeft.HasValue && _settings.WindowTop.HasValue)
            {
                double left = _settings.WindowLeft.Value;
                double top = _settings.WindowTop.Value;

                // デスクトップの作業領域を取得
                double screenWidth = SystemParameters.VirtualScreenWidth;
                double screenHeight = SystemParameters.VirtualScreenHeight;
                double screenLeft = SystemParameters.VirtualScreenLeft;
                double screenTop = SystemParameters.VirtualScreenTop;

                // ウィンドウがデスクトップ外にある場合は左上に寄せる
                if (left < screenLeft || left + Width > screenLeft + screenWidth)
                {
                    left = screenLeft + 50; // 左端から少し離す
                }

                if (top < screenTop || top + Height > screenTop + screenHeight)
                {
                    top = screenTop + 50; // 上端から少し離す
                }

                // ウィンドウがデスクトップ内に収まることを確認
                if (left + Width > screenLeft + screenWidth)
                {
                    left = screenLeft + 50;
                }

                if (top + Height > screenTop + screenHeight)
                {
                    top = screenTop + 50;
                }

                Left = left;
                Top = top;
            }
        }

        private void SaveWindowPosition()
        {
            // 最小化や最大化されていない場合のみ位置とサイズを保存
            if (WindowState == WindowState.Normal)
            {
                // NaN や Infinity をチェックして保存
                if (!double.IsNaN(Left) && !double.IsInfinity(Left) &&
                    !double.IsNaN(Top) && !double.IsInfinity(Top) &&
                    !double.IsNaN(Width) && !double.IsInfinity(Width) &&
                    !double.IsNaN(Height) && !double.IsInfinity(Height))
                {
                    _settings.WindowLeft = Left;
                    _settings.WindowTop = Top;
                    _settings.WindowWidth = Width;
                    _settings.WindowHeight = Height;
                }
            }
        }

        private void SaveSettings()
        {
            if (_isLoadingSettings)
            {
                return;
            }

            // 現在の設定を保存
            _settings.SavePath = SavePathTextBox.Text;
            _settings.ImageFormat = FormatComboBox.SelectedIndex;
            _settings.CopyToClipboard = CopyToClipboardCheckBox.IsChecked ?? true;

            // ファイル名テンプレート（空の場合はデフォルト値）
            _settings.FileNameTemplate = string.IsNullOrWhiteSpace(FileNameTemplateComboBox.Text)
                ? "cap_{yyyyMMdd_HHmmss}_{###}"
                : FileNameTemplateComboBox.Text;

            if (_isRegionSet)
            {
                _settings.CaptureRegion = CaptureRegionSettings.FromRectangle(
                    _captureRegion, _captureRegionDpiScaleX, _captureRegionDpiScaleY);
            }

            SaveWindowPosition();
            _settings.Save();
            SaveMemoToCurrentSavePath();
            _lastAppliedSavePath = NormalizeSavePath(SavePathTextBox.Text);
            _lastAppliedMemoPath = GetMemoFilePath();
        }
    }
}
