using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private void AppendImageTagToEditor(string fileName, string fullPath)
        {
            string imageTag = BuildImageTag(fileName, fullPath);

            if (!string.IsNullOrWhiteSpace(MarkdownEditorTextBox.Text) && !MarkdownEditorTextBox.Text.EndsWith(Environment.NewLine))
            {
                MarkdownEditorTextBox.AppendText(Environment.NewLine);
            }

            MarkdownEditorTextBox.AppendText(imageTag + Environment.NewLine + Environment.NewLine);
            MarkdownEditorTextBox.CaretIndex = MarkdownEditorTextBox.Text.Length;
            MarkdownEditorTextBox.ScrollToEnd();
        }

        private void AppendOcrTextToEditor(string fileName, string recognizedText)
        {
            if (!string.IsNullOrWhiteSpace(MarkdownEditorTextBox.Text) && !MarkdownEditorTextBox.Text.EndsWith(Environment.NewLine))
            {
                MarkdownEditorTextBox.AppendText(Environment.NewLine);
            }

            MarkdownEditorTextBox.AppendText($"### OCR: {fileName}{Environment.NewLine}");
            MarkdownEditorTextBox.AppendText(recognizedText.Trim());
            MarkdownEditorTextBox.AppendText(Environment.NewLine + Environment.NewLine);
            MarkdownEditorTextBox.CaretIndex = MarkdownEditorTextBox.Text.Length;
            MarkdownEditorTextBox.ScrollToEnd();
        }

        private string BuildImageTag(string fileName, string fullPath)
        {
            string memoDirectory = System.IO.Path.GetDirectoryName(GetMemoFilePath())
                ?? NormalizeSavePath(SavePathTextBox.Text);

            string relativePath = System.IO.Path.GetRelativePath(memoDirectory, fullPath)
                .Replace('\\', '/');
            string imageFileName = System.IO.Path.GetFileName(fullPath);

            return $"![{imageFileName}]({relativePath})";
        }

        private string GetMemoFilePath(string? savePath = null, string? fileNameTemplate = null)
        {
            string saveDirectory = NormalizeSavePath(savePath ?? SavePathTextBox.Text);
            string? template = fileNameTemplate;
            if (string.IsNullOrWhiteSpace(template))
            {
                if (!string.IsNullOrWhiteSpace(FileNameTemplateComboBox.Text))
                {
                    template = FileNameTemplateComboBox.Text;
                }
                else if (!string.IsNullOrWhiteSpace(_settings?.FileNameTemplate))
                {
                    template = _settings.FileNameTemplate;
                }
                else
                {
                    template = "cap_{yyyyMMdd_HHmmss}_{###}";
                }
            }
            string extension = FormatComboBox.SelectedIndex == 0 ? "png" : "jpg";

            var tempSettings = new AppSettings
            {
                FileNameTemplate = template
            };
            string generatedFileName = tempSettings.GenerateFileName(Math.Max(_captureCount, 1), extension);
            string imagePath = System.IO.Path.Combine(saveDirectory, generatedFileName);
            string imageDirectory = System.IO.Path.GetDirectoryName(imagePath) ?? saveDirectory;

            return System.IO.Path.Combine(imageDirectory, "memo.md");
        }

        private void LoadMemoAndHistoryFromCurrentSavePath()
        {
            string memoPath = GetMemoFilePath();
            string memoDirectory = System.IO.Path.GetDirectoryName(memoPath)
                ?? NormalizeSavePath(SavePathTextBox.Text);

            string memoText = "# 作業メモ\n";
            if (File.Exists(memoPath))
            {
                try
                {
                    memoText = File.ReadAllText(memoPath, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"memo.md 読み込み失敗: {ex.Message}");
                }
            }

            _isLoadingMemo = true;
            try
            {
                MarkdownEditorTextBox.Text = memoText;
            }
            finally
            {
                _isLoadingMemo = false;
            }

            LoadCaptureHistoryFromMemo(memoText, memoDirectory);
            _lastAppliedMemoPath = memoPath;
        }

        private void LoadCaptureHistoryFromMemo(string memoText, string saveDirectory)
        {
            ThumbnailHistoryPanel.Children.Clear();
            ThumbnailImage.Source = null;
            CurrentCaptureFileNameText.Text = "(未選択)";
            _currentCaptureFileName = null;
            _currentCaptureFullPath = null;

            var matches = Regex.Matches(memoText, @"!\[[^\]]*\]\(([^)]+)\)");
            BitmapImage? latestBitmap = null;
            string? latestFileName = null;
            string? latestFullPath = null;

            foreach (Match match in matches)
            {
                string imagePath = match.Groups[1].Value.Trim().Trim('"');
                if (string.IsNullOrWhiteSpace(imagePath))
                {
                    continue;
                }

                string resolvedPath = System.IO.Path.IsPathRooted(imagePath)
                    ? imagePath
                    : System.IO.Path.GetFullPath(System.IO.Path.Combine(saveDirectory, imagePath.Replace('/', System.IO.Path.DirectorySeparatorChar)));

                if (!File.Exists(resolvedPath))
                {
                    continue;
                }

                BitmapImage? bitmapImage = LoadBitmapImageFromFile(resolvedPath);
                if (bitmapImage == null)
                {
                    continue;
                }

                string fileName = System.IO.Path.GetFileName(resolvedPath);
                AddThumbnailToHistory(bitmapImage, fileName, resolvedPath);

                latestBitmap = bitmapImage;
                latestFileName = fileName;
                latestFullPath = resolvedPath;

                if (ThumbnailHistoryPanel.Children.Count >= MAX_HISTORY_ITEMS)
                {
                    break;
                }
            }

            if (latestBitmap != null && latestFileName != null && latestFullPath != null)
            {
                ThumbnailImage.Source = latestBitmap;
                SetCurrentCapture(latestFileName, latestFullPath);
            }
        }

        private void SaveMemoToCurrentSavePath()
        {
            SaveMemoToFile(GetMemoFilePath());
        }

        private void SaveMemoToFile(string memoPath)
        {
            try
            {
                string directory = System.IO.Path.GetDirectoryName(memoPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(memoPath, MarkdownEditorTextBox.Text, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"memo.md 保存失敗: {ex.Message}");
            }
        }

        private void MarkdownEditorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoadingMemo)
            {
                return;
            }

            SaveMemoToCurrentSavePath();
        }

        private void MarkdownEditorContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            MarkdownEditorPasteMenuItem.IsEnabled = Clipboard.ContainsText() || Clipboard.ContainsImage();
        }

        private void MarkdownEditorPasteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage() && !Clipboard.ContainsText())
            {
                HandlePastedImage();
            }
            else
            {
                MarkdownEditorTextBox.Paste();
            }
        }

        private void MarkdownEditorTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.V)
            {
                if (Clipboard.ContainsImage() && !Clipboard.ContainsText())
                {
                    e.Handled = true;
                    HandlePastedImage();
                    return;
                }
            }

            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.S)
            {
                SaveMemoToCurrentSavePath();
                StatusText.Text = "memo.md を保存しました。";
                e.Handled = true;
            }
        }
    }
}
