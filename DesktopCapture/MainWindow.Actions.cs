using System;
using System.Windows;
using System.Windows.Controls;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private void SelectRegionButton_Click(object sender, RoutedEventArgs e)
        {
            // 領域選択ウィンドウを表示
            var regionSelector = new RegionSelectorWindow();
            if (regionSelector.ShowDialog() == true)
            {
                _captureRegion = regionSelector.SelectedRegion;
                _isRegionSet = true;

                // 選択時の DPI スケールを記録
                var ps = PresentationSource.FromVisual(this);
                _captureRegionDpiScaleX = ps?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                _captureRegionDpiScaleY = ps?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

                RegionText.Text = $"X:{_captureRegion.X}, Y:{_captureRegion.Y}, " +
                                  $"W:{_captureRegion.Width}, H:{_captureRegion.Height}";
                CaptureButton.IsEnabled = true;
                StatusText.Text = "準備完了。キャプチャボタンまたはCtrl+Shift+Cでキャプチャできます。";

                // 設定を保存
                SaveSettings();
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            // フォルダ選択ダイアログ
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "保存先フォルダを選択してください",
                SelectedPath = SavePathTextBox.Text
            };

            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SaveMemoToFile(string.IsNullOrWhiteSpace(_lastAppliedMemoPath) ? GetMemoFilePath(_lastAppliedSavePath, FileNameTemplateComboBox.Text) : _lastAppliedMemoPath);
                SavePathTextBox.Text = NormalizeSavePath(folderDialog.SelectedPath);
                LoadMemoAndHistoryFromCurrentSavePath();
                _captureCount = ScanMaxCaptureCount(SavePathTextBox.Text);
                // 設定を保存
                SaveSettings();
            }
        }

        private void SavePathTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isLoadingSettings)
            {
                return;
            }

            string normalizedNewPath = NormalizeSavePath(SavePathTextBox.Text);
            if (string.Equals(normalizedNewPath, _lastAppliedSavePath, StringComparison.OrdinalIgnoreCase))
            {
                SavePathTextBox.Text = normalizedNewPath;
                return;
            }

            SaveMemoToFile(string.IsNullOrWhiteSpace(_lastAppliedMemoPath) ? GetMemoFilePath(_lastAppliedSavePath, FileNameTemplateComboBox.Text) : _lastAppliedMemoPath);
            SavePathTextBox.Text = normalizedNewPath;
            LoadMemoAndHistoryFromCurrentSavePath();
            _captureCount = ScanMaxCaptureCount(SavePathTextBox.Text);
            SaveSettings();
        }

        private void FormatComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 画像形式変更時に設定を保存
            if (_settings != null)
            {
                SaveSettings();
            }
        }

        private void CopyToClipboardCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            // クリップボードコピー設定変更時に保存
            if (_settings != null)
            {
                SaveSettings();
            }
        }

        private void CaptureButton_Click(object sender, RoutedEventArgs e)
        {
            PerformCapture();
        }

        private void InsertLatestImageTagButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentCaptureFileName) || string.IsNullOrEmpty(_currentCaptureFullPath))
            {
                MessageBox.Show("挿入できるキャプチャがありません。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            AppendImageTagToEditor(_currentCaptureFileName, _currentCaptureFullPath);
            SaveMemoToCurrentSavePath();
            StatusText.Text = $"画像タグを挿入しました: {_currentCaptureFileName}";
        }

        private void InsertLatestOcrButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentCaptureFileName) || string.IsNullOrEmpty(_currentCaptureFullPath))
            {
                MessageBox.Show("OCR対象のキャプチャがありません。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!_ocrService.TryExtractText(_currentCaptureFullPath, out string recognizedText, out string errorMessage))
            {
                MessageBox.Show(errorMessage, "OCRエラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                StatusText.Text = "OCRに失敗しました。";
                return;
            }

            if (string.IsNullOrWhiteSpace(recognizedText))
            {
                StatusText.Text = $"OCR結果は空でした: {_currentCaptureFileName}";
                return;
            }

            AppendOcrTextToEditor(_currentCaptureFileName, recognizedText);
            SaveMemoToCurrentSavePath();
            StatusText.Text = $"OCR結果を挿入しました: {_currentCaptureFileName}";
        }

        private void FileNameTemplateComboBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // テキスト変更時は何もしない（頻繁な保存を避けるため）
        }

        private void FileNameTemplateComboBox_DropDownClosed(object sender, EventArgs e)
        {
            ApplyFileNameTemplateChange();
        }

        private void FileNameTemplateComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            ApplyFileNameTemplateChange();
        }

        private void ApplyFileNameTemplateChange()
        {
            string previousMemoPath = string.IsNullOrWhiteSpace(_lastAppliedMemoPath)
                ? GetMemoFilePath(_lastAppliedSavePath, _settings.FileNameTemplate)
                : _lastAppliedMemoPath;

            if (string.IsNullOrWhiteSpace(FileNameTemplateComboBox.Text))
            {
                FileNameTemplateComboBox.Text = "cap_{yyyyMMdd_HHmmss}_{###}";
                StatusText.Text = "ファイル名テンプレートにデフォルト値を設定しました。";
            }

            string currentTemplate = FileNameTemplateComboBox.Text;
            string newMemoPath = GetMemoFilePath(SavePathTextBox.Text, currentTemplate);
            bool templateChanged = !string.Equals(currentTemplate, _settings.FileNameTemplate, StringComparison.Ordinal);
            bool memoPathChanged = !string.Equals(previousMemoPath, newMemoPath, StringComparison.OrdinalIgnoreCase);

            if (!templateChanged && !memoPathChanged)
            {
                return;
            }

            SaveMemoToFile(previousMemoPath);
            LoadMemoAndHistoryFromCurrentSavePath();

            _settings.AddFileNameHistory(currentTemplate);
            SaveSettings();
            UpdateFileNameHistory();
        }

        private void UpdateFileNameHistory()
        {
            // 現在のテキストを保持
            string currentText = FileNameTemplateComboBox.Text;

            // コンボボックスの項目を更新
            FileNameTemplateComboBox.Items.Clear();
            foreach (var history in _settings.FileNameHistory)
            {
                FileNameTemplateComboBox.Items.Add(history);
            }

            // テキストを復元
            FileNameTemplateComboBox.Text = currentText;
        }

        private void TemplateHelpButton_Click(object sender, RoutedEventArgs e)
        {
            var helpWindow = new TemplateHelpWindow
            {
                Owner = this
            };

            helpWindow.ShowDialog();

            // ヘルプウィンドウで選択されたテンプレートがあれば、テキストボックスに設定
            if (!string.IsNullOrEmpty(helpWindow.SelectedTemplate))
            {
                FileNameTemplateComboBox.Text = helpWindow.SelectedTemplate;
                _settings.AddFileNameHistory(helpWindow.SelectedTemplate);
                SaveSettings();
                UpdateFileNameHistory();
                StatusText.Text = "テンプレートを設定しました。";
            }
        }
    }
}
