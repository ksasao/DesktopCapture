using System;
using System.Text;
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

        private void LoadSettings()
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

            // 画像形式を復元
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

            // ファイル名テンプレートを復元（空の場合はデフォルト値）
            if (string.IsNullOrWhiteSpace(_settings.FileNameTemplate))
            {
                FileNameTemplateTextBox.Text = "cap_{yyyyMMdd_HHmmss}_{###}";
            }
            else
            {
                FileNameTemplateTextBox.Text = _settings.FileNameTemplate;
            }
        }

        private void LoadWindowPosition()
        {
            // ウィンドウ位置とサイズを復元
            if (_settings.WindowWidth.HasValue && _settings.WindowHeight.HasValue)
            {
                Width = _settings.WindowWidth.Value;
                Height = _settings.WindowHeight.Value;
                System.Diagnostics.Debug.WriteLine($"ウィンドウサイズ復元: {Width}x{Height}");
            }

            if (_settings.WindowLeft.HasValue && _settings.WindowTop.HasValue)
            {
                double left = _settings.WindowLeft.Value;
                double top = _settings.WindowTop.Value;

                System.Diagnostics.Debug.WriteLine($"設定値: Left={left}, Top={top}");

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
                System.Diagnostics.Debug.WriteLine($"ウィンドウ位置復元: Left={Left}, Top={Top}");
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
                    System.Diagnostics.Debug.WriteLine($"ウィンドウ位置保存: Left={Left}, Top={Top}, Width={Width}, Height={Height}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"無効な値のため保存スキップ: Left={Left}, Top={Top}, Width={Width}, Height={Height}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"ウィンドウ状態が Normal ではないため保存スキップ: {WindowState}");
            }
        }

        private void SaveSettings()
        {
            // 現在の設定を保存
            _settings.SavePath = SavePathTextBox.Text;
            _settings.ImageFormat = FormatComboBox.SelectedIndex;
            _settings.CopyToClipboard = CopyToClipboardCheckBox.IsChecked ?? true;
            
            // ファイル名テンプレート（空の場合はデフォルト値）
            _settings.FileNameTemplate = string.IsNullOrWhiteSpace(FileNameTemplateTextBox.Text) 
                ? "cap_{yyyyMMdd_HHmmss}_{###}" 
                : FileNameTemplateTextBox.Text;

            if (_isRegionSet)
            {
                _settings.CaptureRegion = CaptureRegionSettings.FromRectangle(_captureRegion);
            }

            SaveWindowPosition();
            _settings.Save();
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
                SavePathTextBox.Text = folderDialog.SelectedPath;
                // 設定を保存
                SaveSettings();
            }
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
                        DisplayThumbnail(bitmap, fileName);
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

        private void DisplayThumbnail(Bitmap bitmap, string fileName)
        {
            // Bitmap を BitmapImage に変換
            BitmapImage bitmapImage = BitmapToBitmapImage(bitmap);
            
            // 最新のキャプチャを大きく表示
            ThumbnailImage.Source = bitmapImage;
            
            // 履歴に追加
            AddThumbnailToHistory(bitmapImage, fileName);
        }

        private void AddThumbnailToHistory(BitmapImage bitmapImage, string fileName)
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
            SaveSettings();
            base.OnClosing(e);
        }

        private void FileNameTemplateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // ファイル名テンプレート変更時に設定を保存
            if (_settings != null)
            {
                SaveSettings();
            }
        }

        private void FileNameTemplateTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // テキストボックスが空の場合、デフォルト値を設定
            if (string.IsNullOrWhiteSpace(FileNameTemplateTextBox.Text))
            {
                FileNameTemplateTextBox.Text = "cap_{yyyyMMdd_HHmmss}_{###}";
                StatusText.Text = "ファイル名テンプレートにデフォルト値を設定しました。";
            }
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
                FileNameTemplateTextBox.Text = helpWindow.SelectedTemplate;
                StatusText.Text = "テンプレートを設定しました。";
            }
        }
    }
}