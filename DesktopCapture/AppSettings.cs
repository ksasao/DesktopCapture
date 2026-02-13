using System;
using System.Collections.Generic;
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
        public List<string> FileNameHistory { get; set; } = new List<string>();
        
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
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        AllowTrailingCommas = true,
                        ReadCommentHandling = JsonCommentHandling.Skip
                    };
                    var settings = JsonSerializer.Deserialize<AppSettings>(json, options);
                    return settings ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                // 読み込みエラーの場合は新しい設定を返す
                // 壊れたファイルをバックアップ
                try
                {
                    if (File.Exists(SettingsFilePath))
                    {
                        File.Move(SettingsFilePath, SettingsFilePath + ".bak", true);
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
            }
            catch (Exception ex)
            {
                // 保存エラーは無視
            }
        }

        /// <summary>
        /// ファイル名履歴に追加（最大10件まで保持）
        /// </summary>
        public void AddFileNameHistory(string fileNameTemplate)
        {
            if (string.IsNullOrWhiteSpace(fileNameTemplate))
                return;

            // 既存の同じ項目を削除
            FileNameHistory.Remove(fileNameTemplate);

            // 先頭に追加
            FileNameHistory.Insert(0, fileNameTemplate);

            // 最大10件まで保持
            if (FileNameHistory.Count > 10)
            {
                FileNameHistory = FileNameHistory.Take(10).ToList();
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
