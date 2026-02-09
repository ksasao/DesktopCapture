using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Linq;

namespace DesktopCapture
{
    public class AppSettings
    {
        public string SavePath { get; set; } = string.Empty;
        public int ImageFormat { get; set; } = 0; // 0: PNG, 1: JPEG
        public bool CopyToClipboard { get; set; } = true;
        public CaptureRegionSettings? CaptureRegion { get; set; }
        public string FileNameTemplate { get; set; } = "cap_{yyyyMMdd_HHmmss}_{###}";
        
        // ウィンドウ位置とサイズ
        public double? WindowLeft { get; set; }
        public double? WindowTop { get; set; }
        public double? WindowWidth { get; set; }
        public double? WindowHeight { get; set; }

        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DesktopCapture",
            "settings.json"
        );

        public static AppSettings Load()
        {
            System.Diagnostics.Debug.WriteLine($"設定ファイルパス: {SettingsFilePath}");
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"設定ファイルが存在します。読み込み中...");
                    string json = File.ReadAllText(SettingsFilePath);
                    System.Diagnostics.Debug.WriteLine($"設定ファイル内容: {json}");
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        AllowTrailingCommas = true,
                        ReadCommentHandling = JsonCommentHandling.Skip
                    };
                    var settings = JsonSerializer.Deserialize<AppSettings>(json, options);
                    System.Diagnostics.Debug.WriteLine($"設定ファイル読み込み成功");
                    return settings ?? new AppSettings();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"設定ファイルが存在しません。新規作成します。");
                }
            }
            catch (Exception ex)
            {
                // 読み込みエラーの場合は新しい設定を返す
                System.Diagnostics.Debug.WriteLine($"設定ファイル読み込みエラー: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"スタックトレース: {ex.StackTrace}");
                // 壊れたファイルをバックアップ
                try
                {
                    if (File.Exists(SettingsFilePath))
                    {
                        File.Move(SettingsFilePath, SettingsFilePath + ".bak", true);
                        System.Diagnostics.Debug.WriteLine($"壊れた設定ファイルを {SettingsFilePath}.bak にバックアップしました");
                    }
                }
                catch { }
            }
            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                string directory = Path.GetDirectoryName(SettingsFilePath)!;
                Directory.CreateDirectory(directory);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.Never
                };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsFilePath, json);
                System.Diagnostics.Debug.WriteLine($"設定ファイル保存成功: {SettingsFilePath}");
            }
            catch (Exception ex)
            {
                // 保存エラーをログに出力
                System.Diagnostics.Debug.WriteLine($"設定ファイル保存エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// テンプレートから実際のファイルパスを生成
        /// </summary>
        public string GenerateFileName(int captureCount, string extension)
        {
            string template = string.IsNullOrEmpty(FileNameTemplate) 
                ? "cap_{yyyyMMdd_HHmmss}_{###}" 
                : FileNameTemplate;

            // 中括弧 {} 内のパターンのみを処理
            string result = System.Text.RegularExpressions.Regex.Replace(template, @"\{([^}]+)\}", match =>
            {
                string pattern = match.Groups[1].Value;
                
                // カウンター（#の連続）の場合
                if (pattern.All(c => c == '#'))
                {
                    int hashCount = pattern.Length;
                    string counterFormat = new string('0', hashCount);
                    return captureCount.ToString(counterFormat);
                }
                
                // 日付時刻パターンの場合
                try
                {
                    return DateTime.Now.ToString(pattern);
                }
                catch
                {
                    // 無効なフォーマットの場合はそのまま返す
                    return match.Value;
                }
            });

            // 拡張子を追加（テンプレートに含まれていない場合）
            if (!result.EndsWith($".{extension}", StringComparison.OrdinalIgnoreCase))
            {
                result += $".{extension}";
            }

            return result;
        }
    }

    public class CaptureRegionSettings
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public Rectangle ToRectangle()
        {
            return new Rectangle(X, Y, Width, Height);
        }

        public static CaptureRegionSettings FromRectangle(Rectangle rect)
        {
            return new CaptureRegionSettings
            {
                X = rect.X,
                Y = rect.Y,
                Width = rect.Width,
                Height = rect.Height
            };
        }
    }
}
