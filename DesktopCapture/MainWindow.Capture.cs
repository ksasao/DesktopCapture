using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DrawingRectangle = System.Drawing.Rectangle;
using DrawingPoint = System.Drawing.Point;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private void PerformCapture()
        {
            if (!_isRegionSet)
            {
                MessageBox.Show("キャプチャ領域が設定されていません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // スクリーンショットを取得
                using (var bitmap = new Bitmap(_captureRegion.Width, _captureRegion.Height))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(_captureRegion.Location,
                            DrawingPoint.Empty, _captureRegion.Size);
                    }

                    // ファイル名を生成
                    _captureCount++;
                    string extension = FormatComboBox.SelectedIndex == 0 ? "png" : "jpg";
                    string fileName = _settings.GenerateFileName(_captureCount, extension);
                    string fullPath = System.IO.Path.Combine(SavePathTextBox.Text, fileName);

                    // ディレクトリが存在しない場合は作成（テンプレートにサブディレクトリが含まれる場合に対応）
                    string? directory = System.IO.Path.GetDirectoryName(fullPath);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // 画像を保存
                    ImageFormat format = FormatComboBox.SelectedIndex == 0
                        ? ImageFormat.Png
                        : ImageFormat.Jpeg;

                    bitmap.Save(fullPath, format);

                    // クリップボードにコピー
                    bool copiedToClipboard = false;
                    if (CopyToClipboardCheckBox.IsChecked == true)
                    {
                        try
                        {
                            BitmapImage bitmapImage = BitmapToBitmapImage(bitmap);
                            Clipboard.SetImage(bitmapImage);
                            copiedToClipboard = true;
                        }
                        catch (Exception ex)
                        {
                            // クリップボードコピー失敗は警告のみ
                            System.Diagnostics.Debug.WriteLine($"クリップボードコピー失敗: {ex.Message}");
                        }
                    }

                    // サムネイルを表示
                    Dispatcher.Invoke(() =>
                    {
                        DisplayThumbnail(bitmap, fileName, fullPath);
                        string statusMessage = $"保存完了: {fileName} (#{_captureCount})";
                        if (copiedToClipboard)
                        {
                            statusMessage += " [クリップボードにコピー済み]";
                        }
                        StatusText.Text = statusMessage;
                        TriggerAutoOcrIfEnabled(fileName, fullPath);

                    });
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"キャプチャに失敗しました: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "キャプチャに失敗しました。";
                });
            }
        }

        private void DisplayThumbnail(Bitmap bitmap, string fileName, string fullPath)
        {
            // Bitmap を BitmapImage に変換
            BitmapImage bitmapImage = BitmapToBitmapImage(bitmap);

            // 最新のキャプチャを大きく表示
            ThumbnailImage.Source = bitmapImage;

            SetCurrentCapture(fileName, fullPath);
            AppendImageTagToEditor(fileName, fullPath);

            // 履歴に追加
            AddThumbnailToHistory(bitmapImage, fileName, fullPath);
        }

        private void AddThumbnailToHistory(BitmapImage bitmapImage, string fileName, string fullPath)
        {
            // サムネイル用のコンテナを作成
            var stackPanel = new StackPanel
            {
                Margin = new Thickness(5),
                Width = 120
            };

            // サムネイル画像
            var image = new System.Windows.Controls.Image
            {
                Source = bitmapImage,
                Width = 100,
                Height = 100,
                Stretch = Stretch.Uniform,
                Margin = new Thickness(0, 0, 0, 5)
            };

            // ファイル名ラベル
            var label = new TextBlock
            {
                Text = fileName,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 10,
                TextAlignment = TextAlignment.Center
            };

            // クリックで拡大表示
            image.MouseLeftButtonDown += (s, e) =>
            {
                ThumbnailImage.Source = bitmapImage;
                SetCurrentCapture(fileName, fullPath);
            };
            image.Cursor = Cursors.Hand;

            stackPanel.Children.Add(image);
            stackPanel.Children.Add(label);

            // 履歴の先頭に追加
            ThumbnailHistoryPanel.Children.Insert(0, stackPanel);

            // 最大数を超えたら古いものを削除
            while (ThumbnailHistoryPanel.Children.Count > MAX_HISTORY_ITEMS)
            {
                ThumbnailHistoryPanel.Children.RemoveAt(ThumbnailHistoryPanel.Children.Count - 1);
            }
        }

        private void SetCurrentCapture(string fileName, string fullPath)
        {
            _currentCaptureFileName = fileName;
            _currentCaptureFullPath = fullPath;
            CurrentCaptureFileNameText.Text = $"({fileName})";
        }

        private void HandlePastedImage()
        {
            try
            {
                BitmapSource? bitmapSource = Clipboard.GetImage();
                if (bitmapSource == null)
                {
                    return;
                }

                _captureCount++;
                string extension = FormatComboBox.SelectedIndex == 0 ? "png" : "jpg";
                string fileName = _settings.GenerateFileName(_captureCount, extension);
                string fullPath = System.IO.Path.Combine(SavePathTextBox.Text, fileName);

                string? directory = System.IO.Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                ImageFormat format = FormatComboBox.SelectedIndex == 0 ? ImageFormat.Png : ImageFormat.Jpeg;

                using Bitmap bitmap = BitmapSourceToBitmap(bitmapSource);
                bitmap.Save(fullPath, format);

                DisplayThumbnail(bitmap, fileName, fullPath);
                StatusText.Text = $"保存完了: {fileName} (#{_captureCount}) [貼り付け]";
                TriggerAutoOcrIfEnabled(fileName, fullPath);  // ← 追加
            }
            catch (Exception ex)
            {
                MessageBox.Show($"画像の保存に失敗しました: {ex.Message}", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "画像の貼り付けに失敗しました。";
            }
        }

        private static Bitmap BitmapSourceToBitmap(BitmapSource bitmapSource)
        {
            var converted = new FormatConvertedBitmap(bitmapSource, PixelFormats.Bgra32, null, 0);
            int width = converted.PixelWidth;
            int height = converted.PixelHeight;
            int stride = width * 4;
            byte[] pixels = new byte[height * stride];
            converted.CopyPixels(pixels, stride, 0);

            var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            BitmapData bitmapData = bitmap.LockBits(
                new DrawingRectangle(0, 0, width, height),
                ImageLockMode.WriteOnly,
                bitmap.PixelFormat);
            Marshal.Copy(pixels, 0, bitmapData.Scan0, pixels.Length);
            bitmap.UnlockBits(bitmapData);
            return bitmap;
        }

        private static BitmapImage BitmapToBitmapImage(Bitmap bitmap)
        {
            using MemoryStream memory = new MemoryStream();
            bitmap.Save(memory, ImageFormat.Png);
            memory.Position = 0;

            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = memory;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();

            return bitmapImage;
        }

        private static BitmapImage? LoadBitmapImageFromFile(string filePath)
        {
            try
            {
                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.UriSource = new Uri(filePath, UriKind.Absolute);
                bitmapImage.EndInit();
                bitmapImage.Freeze();
                return bitmapImage;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"画像読み込み失敗: {filePath} - {ex.Message}");
                return null;
            }
        }

        private int ScanMaxCaptureCount(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return 0;
            string template = _settings?.FileNameTemplate ?? "cap_{yyyyMMdd_HHmmss}_{###}";
            var regex = BuildCounterRegex(template);
            if (regex == null)
                return 0;
            int max = 0;
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                string relative = Path.GetRelativePath(folderPath, file).Replace('\\', '/');
                var m = regex.Match(relative);
                if (m.Success && m.Groups["counter"].Success &&
                    int.TryParse(m.Groups["counter"].Value, out int n) && n > max)
                    max = n;
            }
            return max;
        }

        private static Regex? BuildCounterRegex(string template)
        {
            if (!Regex.IsMatch(template, @"\{#+\}"))
                return null;
            var sb = new StringBuilder("^");
            int i = 0;
            while (i < template.Length)
            {
                if (template[i] == '{')
                {
                    int end = template.IndexOf('}', i + 1);
                    if (end > i)
                    {
                        string inner = template.Substring(i + 1, end - i - 1);
                        sb.Append(inner.All(c => c == '#')
                            ? $@"(?<counter>\d{{{inner.Length},}})"
                            : @"[^/\\]+?");
                        i = end + 1;
                        continue;
                    }
                }
                sb.Append(Regex.Escape(template[i].ToString()));
                i++;
            }
            string built = sb.ToString();
            if (!Regex.IsMatch(built, @"\\\.[a-zA-Z]{2,5}$"))
                sb.Append(@"\.[^./\\]+");
            sb.Append('$');
            return new Regex(sb.ToString(), RegexOptions.IgnoreCase);
        }
    }
}
