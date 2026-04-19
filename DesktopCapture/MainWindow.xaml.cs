using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
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

        private static string NormalizeSavePath(string? savePath)
        {
            return string.IsNullOrWhiteSpace(savePath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                : savePath.Trim();
        }
    }
}
