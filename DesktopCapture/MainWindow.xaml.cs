using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using DrawingRectangle = System.Drawing.Rectangle;

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
        private readonly SnippingToolOcrService _ocrService;
        private double _captureRegionDpiScaleX = 1.0;
        private double _captureRegionDpiScaleY = 1.0;

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
            _ocrService = new SnippingToolOcrService();

            if (!_ocrService.IsAvailable)
            {
                InsertLatestOcrButton.ToolTip = _ocrService.UnavailableReason;
                StatusText.Text = $"OCR未使用: {_ocrService.UnavailableReason}";
            }

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

            DataObject.AddPastingHandler(MarkdownEditorTextBox, OnMarkdownEditorPasting);
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

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
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

        private static string NormalizeSavePath(string? savePath)
        {
            return string.IsNullOrWhiteSpace(savePath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                : savePath.Trim();
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
