using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using Microsoft.Win32;
using DrawingRectangle = System.Drawing.Rectangle;
using DrawingPoint = System.Drawing.Point;
using DrawingSize = System.Drawing.Size;

namespace DesktopCapture
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DrawingRectangle _captureRegion;
        private bool _isRegionSet = false;
        private int _captureCount = 0;
        private const int HOTKEY_ID = 9000;
        private const int MAX_HISTORY_ITEMS = 10;
        private AppSettings _settings;
        private string? _currentCaptureFileName;
        private string? _currentCaptureFullPath;
        private bool _isLoadingMemo;
        private bool _isLoadingSettings;
        private string _lastAppliedSavePath = string.Empty;
        private string _lastAppliedMemoPath = string.Empty;

        // Windows API for global hotkey
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint VK_C = 0x43;

        private System.Windows.Interop.HwndSource? _source;

        public MainWindow()
        {
            InitializeComponent();
            InitializeNoteEditor();
            
            // タイトルにビルド番号を追加
            SetWindowTitle();
            
            // 設定を読み込み
            _settings = AppSettings.Load();
            LoadSettings();
            LoadWindowPosition();
            
            Loaded += MainWindow_Loaded;
            Closed += MainWindow_Closed;
            LocationChanged += MainWindow_LocationChanged;
            SizeChanged += MainWindow_SizeChanged;
        }

        private void SetWindowTitle()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            if (version != null)
            {
                Title = $"デスクトップキャプチャ - v{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
            }
        }

        private void InitializeNoteEditor()
        {
            if (string.IsNullOrWhiteSpace(MarkdownEditorTextBox.Text))
            {
                MarkdownEditorTextBox.Text = "# 作業メモ\n";
            }

            CurrentCaptureFileNameText.Text = "(未選択)";
        }

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
                    _isRegionSet = true;
                    _captureCount = 0;

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
                _settings.CaptureRegion = CaptureRegionSettings.FromRectangle(_captureRegion);
            }

            SaveWindowPosition();
            _settings.Save();
            SaveMemoToCurrentSavePath();
            _lastAppliedSavePath = NormalizeSavePath(SavePathTextBox.Text);
            _lastAppliedMemoPath = GetMemoFilePath();
        }

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

        private void SelectRegionButton_Click(object sender, RoutedEventArgs e)
        {
            // 領域選択ウィンドウを表示
            var regionSelector = new RegionSelectorWindow();
            if (regionSelector.ShowDialog() == true)
            {
                _captureRegion = regionSelector.SelectedRegion;
                _isRegionSet = true;
                _captureCount = 0; // カウンターをリセット
                
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
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "保存先フォルダを選択",
                FileName = "フォルダ選択",
                Filter = "Folder|*.none"
            };

            // フォルダ選択ダイアログの代替として
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

        private string BuildImageTag(string fileName, string fullPath)
        {
            string memoDirectory = System.IO.Path.GetDirectoryName(GetMemoFilePath())
                ?? NormalizeSavePath(SavePathTextBox.Text);

            string relativePath = System.IO.Path.GetRelativePath(memoDirectory, fullPath)
                .Replace('\\', '/');
            string imageFileName = System.IO.Path.GetFileName(fullPath);

            return $"![{imageFileName}]({relativePath})";
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

        private static string NormalizeSavePath(string? savePath)
        {
            return string.IsNullOrWhiteSpace(savePath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                : savePath.Trim();
        }

        private string GetMemoFilePath(string? savePath = null, string? fileNameTemplate = null)
        {
            string saveDirectory = NormalizeSavePath(savePath ?? SavePathTextBox.Text);
            string template = fileNameTemplate;
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

        private BitmapImage? LoadBitmapImageFromFile(string filePath)
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

        private void MarkdownEditorTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.S)
            {
                SaveMemoToCurrentSavePath();
                StatusText.Text = "memo.md を保存しました。";
                e.Handled = true;
            }
        }

        private BitmapImage BitmapToBitmapImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
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
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // ウィンドウを閉じる前に設定を保存
            SaveMemoToCurrentSavePath();
            SaveSettings();
            base.OnClosing(e);
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