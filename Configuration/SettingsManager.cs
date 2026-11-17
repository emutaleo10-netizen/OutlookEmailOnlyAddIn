using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace OutlookEmailOnlyAddIn.Configuration
{
    /// <summary>
    /// ユーザーごとの設定ファイル（%AppData%\OutlookEmailOnlyAddIn\settings.json）の読み書きを管理するクラス
    /// </summary>
    public static class SettingsManager
    {
        private static readonly object SyncLock = new object();
        private static AddInSettings _currentSettings;

        /// <summary>
        /// アプリケーションデータ保存先フォルダパスを取得します。
        /// </summary>
        public static string AppDataDirectory
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, "OutlookEmailOnlyAddIn");
            }
        }

        /// <summary>
        /// 設定ファイルのフルパスを取得します。
        /// </summary>
        public static string SettingsFilePath => Path.Combine(AppDataDirectory, "settings.json");

        /// <summary>
        /// 現在読み込まれている設定を取得します。
        /// </summary>
        public static AddInSettings Current
        {
            get
            {
                lock (SyncLock)
                {
                    if (_currentSettings == null)
                    {
                        _currentSettings = LoadInternal();
                    }
                    return _currentSettings;
                }
            }
        }

        /// <summary>
        /// 設定変更時に発火するイベント
        /// </summary>
        public static event EventHandler SettingsChanged;

        /// <summary>
        /// 設定をディスクから再読み込みします。
        /// </summary>
        public static AddInSettings Reload()
        {
            lock (SyncLock)
            {
                _currentSettings = LoadInternal();
                return _currentSettings;
            }
        }

        /// <summary>
        /// 設定をディスクへ保存し、現在値を更新します。
        /// </summary>
        public static void Save(AddInSettings newSettings)
        {
            if (newSettings == null)
            {
                throw new ArgumentNullException(nameof(newSettings));
            }

            NormalizeDomains(newSettings);

            lock (SyncLock)
            {
                if (!Directory.Exists(AppDataDirectory))
                {
                    Directory.CreateDirectory(AppDataDirectory);
                }

                var serializer = new DataContractJsonSerializer(typeof(AddInSettings));
                using (var ms = new MemoryStream())
                {
                    serializer.WriteObject(ms, newSettings);
                    string json = Encoding.UTF8.GetString(ms.ToArray());
                    File.WriteAllText(SettingsFilePath, json, Encoding.UTF8);
                }

                _currentSettings = newSettings.Clone();
            }

            SettingsChanged?.Invoke(null, EventArgs.Empty);
        }

        private static AddInSettings LoadInternal()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath, Encoding.UTF8);
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                    {
                        var serializer = new DataContractJsonSerializer(typeof(AddInSettings));
                        var loaded = serializer.ReadObject(ms) as AddInSettings;
                        if (loaded != null)
                        {
                            NormalizeDomains(loaded);
                            return loaded;
                        }
                    }
                }
            }
            catch
            {
                // 設定ファイルの破損やアクセスエラー時は既定値にフォールバック
            }

            var defaultSettings = new AddInSettings();
            NormalizeDomains(defaultSettings);
            return defaultSettings;
        }

        /// <summary>
        /// ドメイン一覧の空白除去・先頭@の除去・重複排除を行います。
        /// </summary>
        private static void NormalizeDomains(AddInSettings settings)
        {
            if (settings.ExcludedDomains == null)
            {
                settings.ExcludedDomains = new List<string> { "kinki-sharyo.local" };
                return;
            }

            var normalized = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string raw in settings.ExcludedDomains)
            {
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                string cleaned = raw.Trim().TrimStart('@').TrimStart('.');
                if (!string.IsNullOrEmpty(cleaned) && seen.Add(cleaned))
                {
                    normalized.Add(cleaned);
                }
            }

            settings.ExcludedDomains = normalized;

            if (settings.PeriodicScanIntervalMinutes < 1)
            {
                settings.PeriodicScanIntervalMinutes = 5;
            }
            if (settings.AutoScanLookbackDays < 1)
            {
                settings.AutoScanLookbackDays = 7;
            }
        }
    }
}
