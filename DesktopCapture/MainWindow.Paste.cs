using System;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;

namespace DesktopCapture
{
    public partial class MainWindow : Window
    {
        private static readonly Regex MarkdownUrlRegex = new Regex(
            @"\bhttps?://[^\s\r\n""'<>\[\]()]+",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex HtmlAnchorRegex = new Regex(
            @"<a\s[^>]*\bhref\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]*))[^>]*>(.*?)</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        private static readonly Regex HtmlBlockElementRegex = new Regex(
            @"</?(?:p|div|br|h[1-6]|li|tr|td|th|blockquote|pre|hr|ul|ol|table|article|section|header|footer|main|aside|figure|figcaption)\b[^>]*>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex HtmlTagRegex = new Regex(
            @"<[^>]+>",
            RegexOptions.Compiled);

        private static readonly Regex BareUrlInMarkdownTextRegex = new Regex(
            @"(?<!\]\()\bhttps?://[^\s\r\n""'<>\[\]()]+",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex MarkdownLinkSyntaxRegex = new Regex(
            @"!?\[[^\]]+\]\([^)]+\)",
            RegexOptions.Compiled);

        private void OnMarkdownEditorPasting(object sender, DataObjectPastingEventArgs e)
        {
            string? plainText = e.DataObject.GetData(DataFormats.UnicodeText) as string
                             ?? e.DataObject.GetData(DataFormats.Text) as string;

            if (string.IsNullOrEmpty(plainText))
            {
                return;
            }

            var textBox = MarkdownEditorTextBox;
            int selectionStart = textBox.SelectionStart;
            int selectionLength = textBox.SelectionLength;
            string selectedText = textBox.SelectedText;

            string? replacement = null;

            // HTML形式が利用可能な場合はHTMLからテキストを生成する
            if (e.DataObject.GetDataPresent(DataFormats.Html))
            {
                string? htmlData = e.DataObject.GetData(DataFormats.Html) as string;
                if (!string.IsNullOrEmpty(htmlData))
                {
                    string? htmlConverted = TryConvertHtmlClipboardToText(htmlData);
                    if (!string.IsNullOrEmpty(htmlConverted))
                    {
                        // HTML変換後のテキストに残っている裸のURLをMarkdownリンクに変換する
                        string pasteText = BareUrlInMarkdownTextRegex.IsMatch(htmlConverted)
                            ? BareUrlInMarkdownTextRegex.Replace(htmlConverted, m => $"[{m.Value}]({m.Value})")
                            : htmlConverted;

                        // プレーンテキストが純粋なURLで、かつ選択テキストがある場合は選択テキストをリンクテキストにする
                        string plainTrimmed = plainText.Trim();
                        var plainUrlMatch = MarkdownUrlRegex.Match(plainTrimmed);
                        bool plainIsPureUrl = plainUrlMatch.Success
                            && plainUrlMatch.Index == 0
                            && plainUrlMatch.Length == plainTrimmed.Length;

                        replacement = plainIsPureUrl && !string.IsNullOrEmpty(selectedText)
                            ? $"[{selectedText}]({plainTrimmed})"
                            : pasteText;
                    }
                }
            }

            // HTML変換が得られなかった場合はプレーンテキストパスで処理する
            if (replacement == null)
            {
                // すでにMarkdownリンク構文 [text](url) / ![alt](url) を含む場合は加工せずそのままペーストする
                if (MarkdownLinkSyntaxRegex.IsMatch(plainText))
                {
                    return;
                }

                string trimmed = plainText.Trim();
                var singleUrlMatch = MarkdownUrlRegex.Match(trimmed);
                bool isPureUrl = singleUrlMatch.Success
                                 && singleUrlMatch.Index == 0
                                 && singleUrlMatch.Length == trimmed.Length;

                if (isPureUrl)
                {
                    if (!string.IsNullOrEmpty(selectedText))
                    {
                        replacement = $"[{selectedText}]({trimmed})";
                    }
                    else
                    {
                        replacement = $"[link]({trimmed})";
                    }
                }
                else if (MarkdownUrlRegex.IsMatch(plainText))
                {
                    replacement = MarkdownUrlRegex.Replace(plainText, m => $"[{m.Value}]({m.Value})");
                }
            }

            if (replacement == null)
            {
                return;
            }

            e.CancelCommand();

            string currentText = textBox.Text;
            textBox.Text = currentText[..selectionStart] + replacement + currentText[(selectionStart + selectionLength)..];
            textBox.CaretIndex = selectionStart + replacement.Length;
        }

        private static string? TryConvertHtmlClipboardToText(string htmlClipboard)
        {
            // <!--StartFragment--> ～ <!--EndFragment--> の範囲を取り出す
            const string startMarker = "<!--StartFragment-->";
            const string endMarker = "<!--EndFragment-->";
            int start = htmlClipboard.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase);
            int end = htmlClipboard.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase);
            string fragment = (start >= 0 && end > start)
                ? htmlClipboard.Substring(start + startMarker.Length, end - start - startMarker.Length)
                : htmlClipboard;

            // <a href="url">text</a> → [text](url)
            string result = HtmlAnchorRegex.Replace(fragment, m =>
            {
                string url = (m.Groups[1].Success ? m.Groups[1].Value
                            : m.Groups[2].Success ? m.Groups[2].Value
                            : m.Groups[3].Value).Trim();

                string innerText = HtmlTagRegex.Replace(m.Groups[4].Value, string.Empty).Trim();
                innerText = WebUtility.HtmlDecode(innerText);

                if (string.IsNullOrWhiteSpace(url) || url.StartsWith('#'))
                {
                    return innerText;
                }

                string label = string.IsNullOrWhiteSpace(innerText) ? url : innerText;
                return $"[{label}]({url})";
            });

            // ブロック要素を改行に変換する
            result = HtmlBlockElementRegex.Replace(result, Environment.NewLine);

            // 残りのタグを除去する
            result = HtmlTagRegex.Replace(result, string.Empty);

            // HTMLエンティティをデコードする
            result = WebUtility.HtmlDecode(result);

            // 空白を正規化する（スペース・タブは1つに、3行以上の空行は2行に）
            result = Regex.Replace(result, @"[ \t]+", " ");
            result = Regex.Replace(result, @"(\r\n|\r|\n){3,}", Environment.NewLine + Environment.NewLine);
            result = result.Trim();

            return string.IsNullOrWhiteSpace(result) ? null : result;
        }
    }
}
