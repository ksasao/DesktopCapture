using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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
                return TryExtractText(image, out recognizedText, out errorMessage);
            }
            catch (Exception ex)
            {
                errorMessage = $"OCR実行エラー: {ex.Message}";
                return false;
            }
        }

        public bool TryExtractText(Bitmap bitmap, out string recognizedText, out string errorMessage)
        {
            recognizedText = string.Empty;
            errorMessage = string.Empty;

            if (!IsAvailable)
            {
                errorMessage = UnavailableReason;
                return false;
            }

            try
            {
                LineData[] lines = RunOcr(bitmap);
                if (lines.Length == 0)
                {
                    recognizedText = string.Empty;
                    return true;
                }

                var builder = new StringBuilder();
                foreach (var line in SortLinesForReadingOrder(lines))
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
                        float centerX = 0;
                        float centerY = 0;
                        float width = 0;
                        float height = 0;

                        result = NativeMethods.GetOcrLineBoundingBox(lineHandle, out IntPtr boundingBoxPtr);
                        if (result == 0 && boundingBoxPtr != IntPtr.Zero)
                        {
                            BoundingBox boundingBox = Marshal.PtrToStructure<BoundingBox>(boundingBoxPtr);
                            float minX = Math.Min(Math.Min(boundingBox.x1, boundingBox.x2), Math.Min(boundingBox.x3, boundingBox.x4));
                            float maxX = Math.Max(Math.Max(boundingBox.x1, boundingBox.x2), Math.Max(boundingBox.x3, boundingBox.x4));
                            float minY = Math.Min(Math.Min(boundingBox.y1, boundingBox.y2), Math.Min(boundingBox.y3, boundingBox.y4));
                            float maxY = Math.Max(Math.Max(boundingBox.y1, boundingBox.y2), Math.Max(boundingBox.y3, boundingBox.y4));
                            width = Math.Max(0, maxX - minX);
                            height = Math.Max(0, maxY - minY);
                            centerX = (minX + maxX) / 2f;
                            centerY = (minY + maxY) / 2f;
                        }

                        lines.Add(new LineData
                        {
                            Text = text,
                            CenterX = centerX,
                            CenterY = centerY,
                            Width = width,
                            Height = height,
                            Left = centerX - (width / 2f),
                            Top = centerY - (height / 2f),
                            Right = centerX + (width / 2f),
                            Bottom = centerY + (height / 2f)
                        });
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

        private static IEnumerable<LineData> SortLinesForReadingOrder(LineData[] lines)
        {
            if (lines.Length <= 1)
            {
                return lines;
            }

            var validLines = GetValidLines(lines);

            if (validLines.Count <= 1)
            {
                return lines;
            }

            const float verticalRatio = 1.3f;
            const float overlapRatio = 0.2f;
            const float horizontalGapRatio = 1.8f;

            var (verticalLines, horizontalLines) = SplitLinesByOrientation(validLines, verticalRatio);

            float medianHorizontalHeight = MedianOrDefault(horizontalLines.Select(x => x.Height), 30f);

            float horizontalGapMax = medianHorizontalHeight * horizontalGapRatio;

            var verticalBlocks = GroupLines(verticalLines, isVertical: true, horizontalGapMax, overlapRatio)
                .OrderByDescending(x => x.Left)
                .ThenBy(x => x.Top)
                .ToList();

            var horizontalBlocks = GroupLines(horizontalLines, isVertical: false, horizontalGapMax, overlapRatio)
                .ToList();
            horizontalBlocks = SortHorizontalBlocksByRows(horizontalBlocks);

            // For Japanese documents, vertical blocks are usually read first from right to left.
            // Horizontal blocks are appended after vertical blocks for mixed-layout robustness.
            return verticalBlocks
                .Concat(horizontalBlocks)
                .SelectMany(x => x.Lines)
                .ToArray();
        }

        private static List<LineData> GetValidLines(IEnumerable<LineData> lines)
        {
            // Keep only lines with meaningful geometry. This matches previous behavior.
            return lines
                .Where(x => x.Width > 0 && x.Height > 0)
                .ToList();
        }

        private static (List<LineData> VerticalLines, List<LineData> HorizontalLines) SplitLinesByOrientation(
            IEnumerable<LineData> lines,
            float verticalRatio)
        {
            // Heuristic: h / w >= verticalRatio => vertical writing candidate.
            var verticalLines = lines.Where(x => x.Height / Math.Max(1f, x.Width) >= verticalRatio).ToList();
            var horizontalLines = lines.Where(x => x.Height / Math.Max(1f, x.Width) < verticalRatio).ToList();
            return (verticalLines, horizontalLines);
        }

        private static List<LineBlock> GroupLines(
            List<LineData> lines,
            bool isVertical,
            float horizontalGapMax,
            float overlapRatio)
        {
            if (lines.Count == 0)
            {
                return new List<LineBlock>();
            }

            int n = lines.Count;
            int[] parent = Enumerable.Range(0, n).ToArray();

            int Find(int x)
            {
                while (parent[x] != x)
                {
                    parent[x] = parent[parent[x]];
                    x = parent[x];
                }

                return x;
            }

            void Union(int a, int b)
            {
                int ra = Find(a);
                int rb = Find(b);
                if (ra != rb)
                {
                    parent[rb] = ra;
                }
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if (AreConnected(lines[i], lines[j], isVertical, horizontalGapMax, overlapRatio))
                    {
                        Union(i, j);
                    }
                }
            }

            var groups = new Dictionary<int, List<LineData>>();
            for (int i = 0; i < n; i++)
            {
                int root = Find(i);
                if (!groups.TryGetValue(root, out var group))
                {
                    group = new List<LineData>();
                    groups[root] = group;
                }

                group.Add(lines[i]);
            }

            var blocks = new List<LineBlock>();
            foreach (var group in groups.Values)
            {
                var ordered = group.OrderBy(x => x.Top).ThenBy(x => x.Left).ToList();

                float left = ordered.Min(x => x.Left);
                float top = ordered.Min(x => x.Top);
                float right = ordered.Max(x => x.Right);
                float bottom = ordered.Max(x => x.Bottom);

                blocks.Add(new LineBlock(ordered, left, top, right, bottom));
            }

            return blocks;
        }

        private static bool AreConnected(
            LineData a,
            LineData b,
            bool isVertical,
            float horizontalGapMax,
            float overlapRatio)
        {
            if (isVertical)
            {
                // Vertical writing: lines in the same column should overlap on X,
                // and their distance is measured along Y.
                float overlapXVertical = Overlap1D(a.Left, a.Right, b.Left, b.Right);
                float baseWidthVertical = Math.Max(1f, Math.Min(a.Width, b.Width));
                if ((overlapXVertical / baseWidthVertical) < overlapRatio)
                {
                    return false;
                }

                float gapYVertical = Gap1D(a.Top, a.Bottom, b.Top, b.Bottom);
                return gapYVertical <= horizontalGapMax;
            }

            float overlapX = Overlap1D(a.Left, a.Right, b.Left, b.Right);
            float baseWidth = Math.Max(1f, Math.Min(a.Width, b.Width));
            if ((overlapX / baseWidth) < overlapRatio)
            {
                return false;
            }

            float gapY = Gap1D(a.Top, a.Bottom, b.Top, b.Bottom);
            return gapY <= horizontalGapMax;
        }

        private static float Overlap1D(float a1, float a2, float b1, float b2)
        {
            return Math.Max(0f, Math.Min(a2, b2) - Math.Max(a1, b1));
        }

        private static float Gap1D(float a1, float a2, float b1, float b2)
        {
            if (a2 < b1)
            {
                return b1 - a2;
            }

            if (b2 < a1)
            {
                return a1 - b2;
            }

            return 0f;
        }

        private static float MedianOrDefault(IEnumerable<float> values, float fallback)
        {
            float[] arr = values.Where(x => x > 0).OrderBy(x => x).ToArray();
            if (arr.Length == 0)
            {
                return fallback;
            }

            int m = arr.Length / 2;
            if (arr.Length % 2 == 1)
            {
                return arr[m];
            }

            return (arr[m - 1] + arr[m]) / 2f;
        }

        private static List<LineBlock> SortHorizontalBlocksByRows(List<LineBlock> blocks)
        {
            if (blocks.Count <= 1)
            {
                return blocks
                    .OrderBy(x => x.Top)
                    .ThenBy(x => x.Left)
                    .ToList();
            }

                // Build table-like rows first, then sort cells left->right inside each row.
            float medianHeight = MedianOrDefault(blocks.Select(x => x.Bottom - x.Top), 30f);
            float rowTolerance = Math.Max(8f, medianHeight * 0.45f);

            var sorted = blocks
                .OrderBy(x => x.Top)
                .ThenBy(x => x.Left)
                .ToList();

            var rows = new List<List<LineBlock>>();
            foreach (var block in sorted)
            {
                int bestRowIndex = -1;
                float bestDistance = float.MaxValue;

                for (int i = 0; i < rows.Count; i++)
                {
                    if (!IsSameHorizontalRow(block, rows[i], rowTolerance))
                    {
                        continue;
                    }

                    float rowCenter = (rows[i].Min(x => x.Top) + rows[i].Max(x => x.Bottom)) * 0.5f;
                    float blockCenter = (block.Top + block.Bottom) * 0.5f;
                    float distance = Math.Abs(blockCenter - rowCenter);
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestRowIndex = i;
                    }
                }

                if (bestRowIndex < 0)
                {
                    rows.Add(new List<LineBlock> { block });
                }
                else
                {
                    rows[bestRowIndex].Add(block);
                }
            }

            var flattened = rows
                .OrderBy(r => r.Min(x => x.Top))
                .ThenBy(r => r.Min(x => x.Left))
                .SelectMany(r => r.OrderBy(x => x.Left).ThenBy(x => x.Top))
                .ToList();

            return flattened;
        }

        private static bool IsSameHorizontalRow(LineBlock block, List<LineBlock> row, float tolerance)
        {
            float rowTop = row.Min(x => x.Top);
            float rowBottom = row.Max(x => x.Bottom);
            float rowHeight = Math.Max(1f, rowBottom - rowTop);

            float blockHeight = Math.Max(1f, block.Bottom - block.Top);
            float overlapY = Overlap1D(block.Top, block.Bottom, rowTop, rowBottom);
            float overlapRatio = overlapY / Math.Max(1f, Math.Min(blockHeight, rowHeight));
            if (overlapRatio >= 0.2f)
            {
                return true;
            }

            float rowCenter = (rowTop + rowBottom) * 0.5f;
            float blockCenter = (block.Top + block.Bottom) * 0.5f;
            return Math.Abs(blockCenter - rowCenter) <= tolerance;
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

        [StructLayout(LayoutKind.Sequential)]
        private struct BoundingBox
        {
            public float x1;
            public float y1;
            public float x2;
            public float y2;
            public float x3;
            public float y3;
            public float x4;
            public float y4;
        }

        private sealed class LineData
        {
            public string Text { get; set; } = string.Empty;
            public float CenterX { get; set; }
            public float CenterY { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Right { get; set; }
            public float Bottom { get; set; }
        }

        private sealed class LineBlock
        {
            public LineBlock(List<LineData> lines, float left, float top, float right, float bottom)
            {
                Lines = lines;
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public List<LineData> Lines { get; }
            public float Left { get; }
            public float Top { get; }
            public float Right { get; }
            public float Bottom { get; }
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

            [DllImport("oneocr.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern long GetOcrLineBoundingBox(long line, out IntPtr boundingBoxPtr);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);
    }
}
