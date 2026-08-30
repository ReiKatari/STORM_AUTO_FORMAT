using System;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace GSheetAutoConverter.Services
{
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string DocumentId { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class GSheetConverterService
    {
        private static readonly HttpClient HttpClient = new HttpClient(new HttpClientHandler
        {
            AllowAutoRedirect = true,
            AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
        })
        {
            Timeout = TimeSpan.FromMinutes(2)
        };

        static GSheetConverterService()
        {
            HttpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
        }

        public string ExtractDocId(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            input = input.Trim();

            // 1. Check if input is a local file path
            if (File.Exists(input))
            {
                try
                {
                    string content = File.ReadAllText(input);

                    // Try JSON parsing
                    try
                    {
                        using var doc = JsonDocument.Parse(content);
                        var root = doc.RootElement;
                        if (root.TryGetProperty("doc_id", out var docIdProp))
                        {
                            string? id = docIdProp.GetString();
                            if (!string.IsNullOrEmpty(id)) return id;
                        }
                        if (root.TryGetProperty("url", out var urlProp))
                        {
                            string? url = urlProp.GetString();
                            if (!string.IsNullOrEmpty(url)) return ExtractDocIdFromUrl(url);
                        }
                    }
                    catch
                    {
                        // Content might be plain text URL or shortcut
                    }

                    return ExtractDocIdFromUrl(content);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error reading local file: {ex.Message}");
                }
            }

            // 2. Direct URL or ID passed
            return ExtractDocIdFromUrl(input);
        }

        private string ExtractDocIdFromUrl(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            // Pattern for standard Google Sheet URLs: /d/{DOC_ID}/...
            var match = Regex.Match(input, @"/d/([a-zA-Z0-9-_]+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Pattern for id parameter: ?id={DOC_ID}
            var idMatch = Regex.Match(input, @"[?&]id=([a-zA-Z0-9-_]+)", RegexOptions.IgnoreCase);
            if (idMatch.Success)
            {
                return idMatch.Groups[1].Value;
            }

            // If it's a raw Doc ID (alphanumeric string without slashes or spaces, length >= 20)
            if (Regex.IsMatch(input, @"^[a-zA-Z0-9-_]{20,}$"))
            {
                return input;
            }

            return string.Empty;
        }

        public string SuggestOutputPath(string gsheetFilePath)
        {
            if (string.IsNullOrWhiteSpace(gsheetFilePath)) return string.Empty;

            try
            {
                if (File.Exists(gsheetFilePath))
                {
                    string dir = Path.GetDirectoryName(gsheetFilePath) ?? string.Empty;
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(gsheetFilePath);
                    return Path.Combine(dir, $"{fileNameWithoutExt}.xlsx");
                }
            }
            catch { }

            return string.Empty;
        }

        public async Task<ConversionResult> ConvertGSheetToXlsxAsync(
            string gsheetInput,
            string targetXlsxPath,
            string? googleAuthCookie = null,
            CancellationToken cancellationToken = default)
        {
            var result = new ConversionResult();

            try
            {
                string docId = ExtractDocId(gsheetInput);
                if (string.IsNullOrEmpty(docId))
                {
                    result.ErrorMessage = "Не удалось извлечь Document ID из файла .gsheet или ссылки. Проверьте правильность пути/ссылки.";
                    return result;
                }

                result.DocumentId = docId;

                // Determine target output file path
                if (string.IsNullOrWhiteSpace(targetXlsxPath))
                {
                    targetXlsxPath = SuggestOutputPath(gsheetInput);
                }

                if (string.IsNullOrWhiteSpace(targetXlsxPath))
                {
                    targetXlsxPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"GoogleSheet_{docId}.xlsx");
                }

                result.OutputPath = targetXlsxPath;

                // Google Sheets export URL for high fidelity .xlsx
                string exportUrl = $"https://docs.google.com/spreadsheets/d/{docId}/export?format=xlsx";

                using var request = new HttpRequestMessage(HttpMethod.Get, exportUrl);

                if (!string.IsNullOrWhiteSpace(googleAuthCookie))
                {
                    request.Headers.Add("Cookie", googleAuthCookie);
                }

                using var response = await HttpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                        response.StatusCode == System.Net.HttpStatusCode.Forbidden ||
                        response.RequestMessage?.RequestUri?.Host.Contains("accounts.google.com") == true)
                    {
                        result.ErrorMessage = "Доступ ограничен (401/403). Включите 'Доступ по ссылке' в Google Таблице или укажите Cookie в настройках.";
                    }
                    else
                    {
                        result.ErrorMessage = $"Сервер Google вернул ошибку: {(int)response.StatusCode} {response.ReasonPhrase}";
                    }
                    return result;
                }

                // Check content type or head to ensure we didn't receive HTML login redirect page
                var contentType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
                if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase))
                {
                    result.ErrorMessage = "Получен HTML-ответ вместо файла Excel. Таблица защищена настройками приватности Google Drive. Включите доступ по ссылке для просмотра.";
                    return result;
                }

                // Atomic write to temporary file first
                string targetDir = Path.GetDirectoryName(targetXlsxPath) ?? AppContext.BaseDirectory;
                Directory.CreateDirectory(targetDir);

                string tempFile = Path.Combine(targetDir, $"~{Guid.NewGuid():N}.tmp");

                using (var stream = await response.Content.ReadAsStreamAsync(cancellationToken))
                using (var fileStream = new FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                {
                    await stream.CopyToAsync(fileStream, cancellationToken);
                }

                var tempFileInfo = new FileInfo(tempFile);
                if (tempFileInfo.Length == 0)
                {
                    if (File.Exists(tempFile)) File.Delete(tempFile);
                    result.ErrorMessage = "Скачанный файл имеет нулевой размер.";
                    return result;
                }

                // COMPLETELY OVERWRITE target file in-place with exact same filename
                try
                {
                    File.Copy(tempFile, targetXlsxPath, overwrite: true);
                }
                catch (IOException ex)
                {
                    result.ErrorMessage = $"Файл Excel заблокирован от перезаписи другой программой (например, открыт в MS Excel): {ex.Message}";
                    return result;
                }
                finally
                {
                    if (File.Exists(tempFile))
                    {
                        try { File.Delete(tempFile); } catch { }
                    }
                }

                result.Success = true;
                result.FileSizeBytes = tempFileInfo.Length;
            }
            catch (OperationCanceledException)
            {
                result.ErrorMessage = "Операция конвертации была отменена.";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Ошибка конвертации: {ex.Message}";
            }

            return result;
        }
    }
}
