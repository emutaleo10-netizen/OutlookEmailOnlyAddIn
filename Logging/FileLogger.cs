using System;
using System.IO;
using System.Text;
using OutlookEmailOnlyAddIn.Configuration;

namespace OutlookEmailOnlyAddIn.Logging
{
    /// <summary>
    /// 日付単位のローテーションと詳細度フィルタを備えたファイルロガー
    /// </summary>
    public static class FileLogger
    {
        private static readonly object WriteLock = new object();

        /// <summary>
        /// ログファイルの保存先ディレクトリパスを取得します。
        /// </summary>
        public static string LogDirectory => Path.Combine(SettingsManager.AppDataDirectory, "logs");

        /// <summary>
        /// デバッグレベルのログを記録します（設定が Debug の場合のみ出力）。
        /// </summary>
        public static void Debug(string message)
        {
            Write(LogLevel.Debug, "DEBUG", message, null);
        }

        /// <summary>
        /// 情報レベルのログを記録します（設定が Info または Debug の場合に出力）。
        /// </summary>
        public static void Info(string message)
        {
            Write(LogLevel.Info, "INFO ", message, null);
        }

        /// <summary>
        /// エラーレベルのログを記録します（常に出力）。
        /// </summary>
        public static void Error(string message, Exception ex = null)
        {
            Write(LogLevel.Error, "ERROR", message, ex);
        }

        private static void Write(LogLevel requiredLevel, string levelTag, string message, Exception ex)
        {
            try
            {
                var settings = SettingsManager.Current;
                if ((int)settings.LogVerbosity < (int)requiredLevel)
                {
                    return;
                }

                lock (WriteLock)
                {
                    if (!Directory.Exists(LogDirectory))
                    {
                        Directory.CreateDirectory(LogDirectory);
                    }

                    string fileName = $"addin-{DateTime.Now:yyyyMMdd}.log";
                    string filePath = Path.Combine(LogDirectory, fileName);

                    var sb = new StringBuilder();
                    sb.AppendFormat("{0:yyyy-MM-dd HH:mm:ss.fff} [{1}] {2}", DateTime.Now, levelTag, message);
                    sb.AppendLine();

                    if (ex != null)
                    {
                        sb.AppendFormat("  Exception: {0}: {1}", ex.GetType().FullName, ex.Message);
                        sb.AppendLine();
                        if (!string.IsNullOrEmpty(ex.StackTrace))
                        {
                            sb.AppendLine(ex.StackTrace);
                        }
                    }

                    File.AppendAllText(filePath, sb.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // ログ出力失敗によって本処理やメール送信を妨げないよう例外を握りつぶす
            }
        }
    }
}
