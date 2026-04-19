using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using DrawingRectangle = System.Drawing.Rectangle;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private System.Windows.Point _thumbnailSelectionStart;
        private bool _isThumbnailSelecting;

        private void ThumbnailCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ThumbnailImage.Source == null) return;

            _thumbnailSelectionStart = e.GetPosition(ThumbnailSelectionCanvas);
            _isThumbnailSelecting = true;

            Canvas.SetLeft(ThumbnailSelectionRect, _thumbnailSelectionStart.X);
            Canvas.SetTop(ThumbnailSelectionRect, _thumbnailSelectionStart.Y);
            ThumbnailSelectionRect.Width = 0;
            ThumbnailSelectionRect.Height = 0;
            ThumbnailSelectionRect.Visibility = Visibility.Visible;

            ThumbnailSelectionCanvas.CaptureMouse();
        }

        private void ThumbnailCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isThumbnailSelecting) return;

            var current = e.GetPosition(ThumbnailSelectionCanvas);
            double x = Math.Min(current.X, _thumbnailSelectionStart.X);
            double y = Math.Min(current.Y, _thumbnailSelectionStart.Y);
            double w = Math.Abs(current.X - _thumbnailSelectionStart.X);
            double h = Math.Abs(current.Y - _thumbnailSelectionStart.Y);

            Canvas.SetLeft(ThumbnailSelectionRect, x);
            Canvas.SetTop(ThumbnailSelectionRect, y);
            ThumbnailSelectionRect.Width = w;
            ThumbnailSelectionRect.Height = h;
        }

        private void ThumbnailCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isThumbnailSelecting) return;

            _isThumbnailSelecting = false;
            ThumbnailSelectionCanvas.ReleaseMouseCapture();
            ThumbnailSelectionRect.Visibility = Visibility.Collapsed;

            double selX = Canvas.GetLeft(ThumbnailSelectionRect);
            double selY = Canvas.GetTop(ThumbnailSelectionRect);
            double selW = ThumbnailSelectionRect.Width;
            double selH = ThumbnailSelectionRect.Height;

            if (selW < 5 || selH < 5) return;

            PerformOcrOnThumbnailRegion(selX, selY, selW, selH);
        }

        /// <summary>
        /// キャプチャ完了後に「自動的にOCR」チェックボックスが有効であれば
        /// 画像全体に対して OCR を実行し Markdown 領域に追記する。
        /// </summary>
        internal void TriggerAutoOcrIfEnabled(string captureFileName, string captureFullPath)
        {
            if (AutoOcrCheckBox.IsChecked != true) return;
            if (!_ocrService.IsAvailable) return;

            StatusText.Text = "自動OCR処理中...";

            Task.Run(() =>
            {
                bool success;
                string recognizedText;
                string errorMessage;

                try
                {
                    success = _ocrService.TryExtractText(captureFullPath, out recognizedText, out errorMessage);
                }
                catch (Exception ex)
                {
                    success = false;
                    recognizedText = string.Empty;
                    errorMessage = ex.Message;
                }

                Dispatcher.Invoke(() =>
                {
                    if (!success || string.IsNullOrWhiteSpace(recognizedText))
                    {
                        StatusText.Text = string.IsNullOrEmpty(errorMessage)
                            ? "自動OCR: テキストが検出されませんでした。"
                            : $"自動OCRエラー: {errorMessage}";
                        return;
                    }

                    AppendOcrTextToEditor(captureFileName, recognizedText);
                    SaveMemoToCurrentSavePath();
                    StatusText.Text = $"自動OCR完了: {recognizedText.Length}文字を検出しました。";
                });
            });
        }

        /// <summary>
        /// サムネイル上でユーザーが選択した矩形領域を元画像ピクセル座標へ変換し、
        /// 当該領域に対して OCR を実行して Markdown エディタへテキストを挿入します。
        /// </summary>
        /// <param name="selX">選択矩形の左上 X（キャンバス座標）</param>
        /// <param name="selY">選択矩形の左上 Y（キャンバス座標）</param>
        /// <param name="selW">選択矩形の幅（キャンバス座標）</param>
        /// <param name="selH">選択矩形の高さ（キャンバス座標）</param>
        private void PerformOcrOnThumbnailRegion(double selX, double selY, double selW, double selH)
        {
            if (!_ocrService.IsAvailable)
            {
                StatusText.Text = $"OCR利用不可: {_ocrService.UnavailableReason}";
                return;
            }

            if (ThumbnailImage.Source == null || string.IsNullOrEmpty(_currentCaptureFullPath))
            {
                StatusText.Text = "OCR対象の画像がありません。";
                return;
            }

            // ThumbnailImage がキャンバス上でどこに描画されているかを取得
            var imageTransform = ThumbnailImage.TransformToVisual(ThumbnailSelectionCanvas);
            var imageOrigin = imageTransform.Transform(new System.Windows.Point(0, 0));

            // Viewbox のスケールを含むキャンバス座標系での実際の描画サイズを算出
            // RenderSize は Viewbox スケール前の自然サイズなので、右下角を変換して差分を取る
            var imageBottomRight = imageTransform.Transform(
                new System.Windows.Point(ThumbnailImage.RenderSize.Width, ThumbnailImage.RenderSize.Height));
            double imageRenderWidth = imageBottomRight.X - imageOrigin.X;
            double imageRenderHeight = imageBottomRight.Y - imageOrigin.Y;

            if (imageRenderWidth <= 0 || imageRenderHeight <= 0) return;

            // 選択範囲を画像の描画領域内にクランプ（すべてキャンバス座標系）
            double clampedLeft = Math.Max(selX, imageOrigin.X);
            double clampedTop = Math.Max(selY, imageOrigin.Y);
            double clampedRight = Math.Min(selX + selW, imageOrigin.X + imageRenderWidth);
            double clampedBottom = Math.Min(selY + selH, imageOrigin.Y + imageRenderHeight);

            if (clampedRight <= clampedLeft || clampedBottom <= clampedTop) return;

            // 画像要素内の相対座標に変換（キャンバス座標系）
            double imgRelX = clampedLeft - imageOrigin.X;
            double imgRelY = clampedTop - imageOrigin.Y;
            double imgRelW = clampedRight - clampedLeft;
            double imgRelH = clampedBottom - clampedTop;

            // 元画像のピクセル座標にスケール変換
            // imageRenderWidth はキャンバス座標系の値なので imgRelX と単位が一致する
            var bitmapSource = (BitmapSource)ThumbnailImage.Source;
            double scaleX = bitmapSource.PixelWidth / imageRenderWidth;
            double scaleY = bitmapSource.PixelHeight / imageRenderHeight;

            int pixX = (int)Math.Round(imgRelX * scaleX);
            int pixY = (int)Math.Round(imgRelY * scaleY);
            int pixW = (int)Math.Round(imgRelW * scaleX);
            int pixH = (int)Math.Round(imgRelH * scaleY);

            // 元画像のサイズ内にクランプ
            pixX = Math.Max(0, Math.Min(pixX, bitmapSource.PixelWidth - 1));
            pixY = Math.Max(0, Math.Min(pixY, bitmapSource.PixelHeight - 1));
            pixW = Math.Max(1, Math.Min(pixW, bitmapSource.PixelWidth - pixX));
            pixH = Math.Max(1, Math.Min(pixH, bitmapSource.PixelHeight - pixY));

            // OCR精度向上のためキャンバスサイズを自動拡張（パディング領域は白で塗りつぶし）
            const int autoPadding = 100;
            int paddedX = Math.Max(0, pixX - autoPadding);
            int paddedY = Math.Max(0, pixY - autoPadding);
            int paddedW = Math.Min(bitmapSource.PixelWidth - paddedX, pixX + pixW + autoPadding - paddedX);
            int paddedH = Math.Min(bitmapSource.PixelHeight - paddedY, pixY + pixH + autoPadding - paddedY);

            StatusText.Text = "OCR処理中...";

            string captureFileName = _currentCaptureFileName ?? string.Empty;
            string captureFullPath = _currentCaptureFullPath;

            Task.Run(() =>
            {
                bool success;
                string recognizedText;
                string errorMessage;

                try
                {
                    const int minOcrSize = 64;

                    // パディングキャンバスのサイズ（最小サイズを保証）
                    int canvasW = Math.Max(paddedW, minOcrSize);
                    int canvasH = Math.Max(paddedH, minOcrSize);

                    // 選択領域の描画位置（autoPadding 分のオフセット + 最小サイズ調整）
                    int drawX = (pixX - paddedX) + (canvasW - paddedW) / 2;
                    int drawY = (pixY - paddedY) + (canvasH - paddedH) / 2;

                    using var originalBitmap = new Bitmap(captureFullPath);

                    // 選択領域のみをクロップし、白背景キャンバスに配置
                    // パディング領域は白で塗りつぶすことで周囲の文字をOCRに含めない
                    var selectedRect = new DrawingRectangle(pixX, pixY, pixW, pixH);
                    using var selectedBitmap = originalBitmap.Clone(selectedRect, originalBitmap.PixelFormat);

                    using var ocrBitmap = new Bitmap(canvasW, canvasH);
                    using var g = Graphics.FromImage(ocrBitmap);
                    g.Clear(Color.White);
                    g.DrawImage(selectedBitmap, drawX, drawY, pixW, pixH);

                    success = _ocrService.TryExtractText(ocrBitmap, out recognizedText, out errorMessage);
                }
                catch (Exception ex)
                {
                    success = false;
                    recognizedText = string.Empty;
                    errorMessage = ex.Message;
                }

                Dispatcher.Invoke(() =>
                {
                    if (!success || string.IsNullOrWhiteSpace(recognizedText))
                    {
                        StatusText.Text = string.IsNullOrEmpty(errorMessage)
                            ? "OCR: テキストが検出されませんでした。"
                            : $"OCRエラー: {errorMessage}";
                        return;
                    }

                    // 選択中テキストがあれば置換、なければカーソル位置へ挿入
                    string insertText = recognizedText.Trim();
                    MarkdownEditorTextBox.SelectedText = insertText;

                    SaveMemoToCurrentSavePath();
                    StatusText.Text = $"OCR完了: {insertText.Length}文字を挿入しました。";
                });

            });
        }
    }
}

