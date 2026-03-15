using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace DesktopCapture
{
    internal sealed class SnippingToolOcrService
    {
        private const string OcrModelKey = "kj)TGtrK>f]b[Piow.gU+nC@s\"\"\"\"\"\"4";
        private const string OneOcrDllFileName = "oneocr.dll";
        private const string OneOcrModelFileName = "oneocr.onemodel";

        private readonly string _installPath = string.Empty;
        private long _context;

        public bool IsAvailable { get; }

        public string UnavailableReason { get; } = "OCRを初期化できませんでした。";

        public SnippingToolOcrService()
        {
            try
            {
                string sourcePath = FindSnippingToolPath();
                if (string.IsNullOrWhiteSpace(sourcePath))
                {
                    UnavailableReason = "Snipping Tool のインストール先が見つかりません。";
                    return;
                }

                _installPath = Path.Combine(AppContext.BaseDirectory, "ocr-runtime");
                if (!TryPrepareRuntimeFiles(sourcePath, _installPath, out string prepareError))
                {
                    UnavailableReason = prepareError;
                    return;
                }

                SetDllDirectory(_installPath);

                long result = NativeMethods.CreateOcrInitOptions(out _context);
                if (result != 0)
                {
                    UnavailableReason = $"OCR初期化オプション作成に失敗しました。Error: {result}";
                    return;
                }

                result = NativeMethods.OcrInitOptionsSetUseModelDelayLoad(_context, 0);
                if (result != 0)
                {
                    UnavailableReason = $"OCR初期化に失敗しました。Error: {result}";
                    return;
                }

                IsAvailable = true;
                UnavailableReason = string.Empty;
            }
            catch (DllNotFoundException)
            {
                UnavailableReason = "oneocr.dll の読み込みに失敗しました。`ocr-runtime` フォルダへのコピーに失敗している可能性があります。";
            }
            catch (BadImageFormatException)
            {
                UnavailableReason = "oneocr.dll のアーキテクチャが一致しません。x64 環境で実行してください。";
            }
            catch (Exception ex)
            {
                UnavailableReason = $"OCR初期化エラー: {ex.Message}";
            }
        }

        private static bool TryPrepareRuntimeFiles(string sourcePath, string runtimePath, out string errorMessage)
        {
            errorMessage = string.Empty;

            string sourceOneOcrDllPath = Path.Combine(sourcePath, OneOcrDllFileName);
            string sourceOneOcrModelPath = Path.Combine(sourcePath, OneOcrModelFileName);
            if (!File.Exists(sourceOneOcrDllPath) || !File.Exists(sourceOneOcrModelPath))
            {
                errorMessage = "Snipping Tool 配下に `oneocr.dll` / `oneocr.onemodel` が見つかりません。";
                return false;
            }

            try
            {
                Directory.CreateDirectory(runtimePath);

                CopyIfNeeded(sourceOneOcrDllPath, Path.Combine(runtimePath, OneOcrDllFileName));
                CopyIfNeeded(sourceOneOcrModelPath, Path.Combine(runtimePath, OneOcrModelFileName));

                foreach (string onnxRuntimeDllPath in Directory.GetFiles(sourcePath, "onnxruntime*.dll", SearchOption.TopDirectoryOnly))
                {
                    string destinationPath = Path.Combine(runtimePath, Path.GetFileName(onnxRuntimeDllPath));
                    CopyIfNeeded(onnxRuntimeDllPath, destinationPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"OCRランタイムの準備に失敗しました: {ex.Message}"
                    + Environment.NewLine
                    + "管理者PowerShellで以下を実行してください:"
                    + Environment.NewLine
                    + BuildManualCopyCommand(sourcePath, runtimePath);
                return false;
            }
        }

        private static void CopyIfNeeded(string sourcePath, string destinationPath)
        {
            if (!File.Exists(destinationPath))
            {
                File.Copy(sourcePath, destinationPath, true);
                return;
            }

            var sourceInfo = new FileInfo(sourcePath);
            var destinationInfo = new FileInfo(destinationPath);
            if (sourceInfo.Length != destinationInfo.Length || sourceInfo.LastWriteTimeUtc > destinationInfo.LastWriteTimeUtc)
            {
                File.Copy(sourcePath, destinationPath, true);
            }
        }

        private static string BuildManualCopyCommand(string sourcePath, string runtimePath)
        {
            string escapedSourcePath = sourcePath.Replace("'", "''");
            string escapedRuntimePath = runtimePath.Replace("'", "''");

            return "$src='" + escapedSourcePath + "'; $dst='" + escapedRuntimePath + "'; New-Item -ItemType Directory -Force -Path $dst | Out-Null; "
                + "Copy-Item (Join-Path $src 'oneocr.dll') $dst -Force; "
                + "Copy-Item (Join-Path $src 'oneocr.onemodel') $dst -Force; "
                + "Copy-Item (Join-Path $src 'onnxruntime*.dll') $dst -Force";
        }

        public bool TryExtractText(string imagePath, out string recognizedText, out string errorMessage)
        {
            recognizedText = string.Empty;
            errorMessage = string.Empty;

            if (!IsAvailable)
            {
                errorMessage = UnavailableReason;
                return false;
            }

            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            {
                errorMessage = "OCR対象の画像ファイルが見つかりません。";
                return false;
            }

            try
            {
                using Bitmap image = new Bitmap(imagePath);
                LineData[] lines = RunOcr(image);
                if (lines.Length == 0)
                {
                    recognizedText = string.Empty;
                    return true;
                }

                var builder = new StringBuilder();
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line.Text))
                    {
                        if (builder.Length > 0)
                        {
                            builder.AppendLine();
                        }

                        builder.Append(line.Text.Trim());
                    }
                }

                recognizedText = builder.ToString();
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"OCR実行エラー: {ex.Message}";
                return false;
            }
        }

        private LineData[] RunOcr(Bitmap sourceImage)
        {
            using Bitmap formattedImage = new Bitmap(sourceImage.Width, sourceImage.Height, PixelFormat.Format32bppArgb);
            using (Graphics graphics = Graphics.FromImage(formattedImage))
            {
                graphics.Clear(Color.White);
                graphics.DrawImage(sourceImage, 0, 0);
            }

            BitmapData bitmapData = formattedImage.LockBits(
                new Rectangle(0, 0, formattedImage.Width, formattedImage.Height),
                ImageLockMode.ReadOnly,
                formattedImage.PixelFormat);

            try
            {
                var imageData = new Img
                {
                    t = 3,
                    col = formattedImage.Width,
                    row = formattedImage.Height,
                    _unk = 0,
                    step = Image.GetPixelFormatSize(formattedImage.PixelFormat) / 8 * formattedImage.Width,
                    data_ptr = bitmapData.Scan0
                };

                string modelPath = Path.Combine(_installPath, "oneocr.onemodel");
                long result = NativeMethods.CreateOcrPipeline(modelPath, OcrModelKey, _context, out long pipeline);
                if (result != 0)
                {
                    throw new InvalidOperationException($"OCRパイプライン作成に失敗しました。Error: {result}");
                }

                result = NativeMethods.CreateOcrProcessOptions(out long options);
                if (result != 0)
                {
                    throw new InvalidOperationException($"OCR処理オプション作成に失敗しました。Error: {result}");
                }

                result = NativeMethods.OcrProcessOptionsSetMaxRecognitionLineCount(options, 1000);
                if (result != 0)
                {
                    throw new InvalidOperationException($"OCR行数設定に失敗しました。Error: {result}");
                }

                result = NativeMethods.RunOcrPipeline(pipeline, ref imageData, options, out long instance);
                if (result != 0)
                {
                    throw new InvalidOperationException($"OCR実行に失敗しました。Error: {result}");
                }

                result = NativeMethods.GetOcrLineCount(instance, out long lineCount);
                if (result != 0)
                {
                    throw new InvalidOperationException($"OCR結果取得に失敗しました。Error: {result}");
                }

                var lines = new List<LineData>();
                for (long i = 0; i < lineCount; i++)
                {
                    result = NativeMethods.GetOcrLine(instance, i, out long lineHandle);
                    if (result != 0 || lineHandle == 0)
                    {
                        continue;
                    }

                    result = NativeMethods.GetOcrLineContent(lineHandle, out IntPtr lineContentPtr);
                    if (result != 0 || lineContentPtr == IntPtr.Zero)
                    {
                        continue;
                    }

                    string? text = PtrToStringUtf8(lineContentPtr);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        lines.Add(new LineData { Text = text });
                    }
                }

                return lines.ToArray();
            }
            finally
            {
                formattedImage.UnlockBits(bitmapData);
            }
        }

        private static string? PtrToStringUtf8(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero)
            {
                return null;
            }

            int length = 0;
            while (Marshal.ReadByte(ptr, length) != 0)
            {
                length++;
            }

            byte[] buffer = new byte[length];
            Marshal.Copy(ptr, buffer, 0, length);
            return Encoding.UTF8.GetString(buffer);
        }

        private static string FindSnippingToolPath()
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -Command \"$pkg=Get-AppxPackage -Name Microsoft.ScreenSketch; if($pkg){ if($pkg.InstallLocation){ Join-Path $pkg.InstallLocation 'SnippingTool' }; if($pkg.InstallLocation -and $pkg.AppInstallPath){ Join-Path (Join-Path $pkg.InstallLocation $pkg.AppInstallPath) 'SnippingTool' } }\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                return string.Empty;
            }

            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            if (string.IsNullOrWhiteSpace(output))
            {
                return string.Empty;
            }

            string[] candidates = output
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string candidate in candidates)
            {
                string normalized = candidate.Trim();
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                string oneOcrPath = Path.Combine(normalized, "oneocr.dll");
                if (Directory.Exists(normalized) && File.Exists(oneOcrPath))
                {
                    return normalized;
                }
            }

            return string.Empty;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Img
        {
            public int t;
            public int col;
            public int row;
            public int _unk;
            public long step;
            public IntPtr data_ptr;
        }

        private sealed class LineData
        {
            public string Text { get; set; } = string.Empty;
        }

        private static class NativeMethods
        {
            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long CreateOcrInitOptions(out long ctx);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long OcrInitOptionsSetUseModelDelayLoad(long ctx, byte flag);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long CreateOcrPipeline(string modelPath, string key, long ctx, out long pipeline);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long CreateOcrProcessOptions(out long options);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long OcrProcessOptionsSetMaxRecognitionLineCount(long options, long count);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long RunOcrPipeline(long pipeline, ref Img img, long options, out long instance);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long GetOcrLineCount(long instance, out long count);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long GetOcrLine(long instance, long index, out long line);

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long GetOcrLineContent(long line, out IntPtr content);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);
    }
}
