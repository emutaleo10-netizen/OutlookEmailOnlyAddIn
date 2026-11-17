using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using OutlookEmailOnlyAddIn.Configuration;
using OutlookEmailOnlyAddIn.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// 受信トレイ・送信済みアイテムの新着監視（Items.ItemAdd + SyncEnd）および
    /// 起動時・定期差分スキャン（ItemAdd の16件制限対策）を管理するクラス
    /// </summary>
    public sealed class FolderMonitor : IDisposable
    {
        private readonly Outlook.Application _application;

        // GC によるイベントハンドラ消失を防ぐため、監視対象の Items と SyncObject をフィールドで強参照保持する
        private readonly List<Outlook.Items> _inboxItemsList = new List<Outlook.Items>();
        private readonly List<Outlook.Items> _sentItemsList = new List<Outlook.Items>();
        private readonly List<Outlook.SyncObject> _syncObjectsList = new List<Outlook.SyncObject>();

        // 起動時遅延スキャン・定期スキャン用タイマー（STA / メイン UI スレッド上で動作）
        private Timer _periodicTimer;
        private Timer _chunkWorkerTimer;

        // バックグラウンド差分スキャン用のキュー（UI 固まりを防ぐため1ティックあたり少数件ずつ処理する）
        private sealed class PendingScanItem
        {
            public string EntryId { get; set; }
            public string StoreId { get; set; }
            public bool IsSentFolder { get; set; }
        }

        private readonly Queue<PendingScanItem> _scanQueue = new Queue<PendingScanItem>();
        private readonly HashSet<string> _queuedEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _isCollectingCatchUp;

        // 1回のタイマーティック（50ms）で処理する最大件数
        private const int ItemsPerTick = 5;

        public FolderMonitor(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// フォルダ監視および定期差分スキャンを開始します。
        /// </summary>
        public void Start()
        {
            HookAllStoresInboxAndSentFolders();
            HookSyncObjects();

            SettingsManager.SettingsChanged += OnSettingsChanged;

            // チャンク処理用ワーカータイマー（50ms 間隔でキュー内のメールを少しずつ処理）
            _chunkWorkerTimer = new Timer { Interval = 50 };
            _chunkWorkerTimer.Tick += ChunkWorkerTimer_Tick;

            // 定期スキャン用タイマー（まずは起動5秒後に初回差分スキャンを実行し、その後設定間隔へ切り替え）
            _periodicTimer = new Timer { Interval = 5000 };
            _periodicTimer.Tick += PeriodicTimer_Tick;
            _periodicTimer.Start();
        }

        /// <summary>
        /// プロファイル内の全 PST ストアの「受信トレイ」と「送信済みアイテム」の ItemAdd イベントを購読します。
        /// </summary>
        private void HookAllStoresInboxAndSentFolders()
        {
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;

            try
            {
                session = _application.Session;
                stores = session.Stores;
                int storeCount = stores.Count;

                for (int i = 1; i <= storeCount; i++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        store = stores[i];
                        HookFolderIfPresent(store, Outlook.OlDefaultFolders.olFolderInbox, isSentFolder: false);
                        HookFolderIfPresent(store, Outlook.OlDefaultFolders.olFolderSentMail, isSentFolder: true);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Debug($"ストア({i})のフォルダ監視登録をスキップしました: {ex.Message}");
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(store);
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("受信トレイ/送信済みアイテムの監視登録中にエラーが発生しました。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(stores);
                MailAddressHelper.ReleaseCom(session);
            }
        }

        private void HookFolderIfPresent(
            Outlook.Store store,
            Outlook.OlDefaultFolders defaultFolderType,
            bool isSentFolder)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = store.GetDefaultFolder(defaultFolderType);
                if (folder == null)
                {
                    return;
                }

                Outlook.Items items = folder.Items;
                if (items == null)
                {
                    return;
                }

                if (isSentFolder)
                {
                    items.ItemAdd += SentItems_ItemAdd;
                    _sentItemsList.Add(items);
                }
                else
                {
                    items.ItemAdd += InboxItems_ItemAdd;
                    _inboxItemsList.Add(items);
                }

                FileLogger.Info($"[監視開始] ストア '{store.DisplayName}' の '{folder.Name}' を監視対象に登録しました。");
            }
            catch
            {
                // 該当ストアにデフォルトフォルダが存在しない場合（アーカイブ専用PST等）は無視
            }
            finally
            {
                MailAddressHelper.ReleaseCom(folder);
            }
        }

        /// <summary>
        /// 送受信グループ（SyncObjects）の完了イベント（SyncEnd）を購読し、
        /// 大量受信で ItemAdd が発火しなかった場合でも送受信完了直後に差分スキャンを実行します。
        /// </summary>
        private void HookSyncObjects()
        {
            Outlook.NameSpace session = null;
            Outlook.SyncObjects syncObjects = null;

            try
            {
                session = _application.Session;
                syncObjects = session.SyncObjects;
                if (syncObjects == null)
                {
                    return;
                }

                int count = syncObjects.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.SyncObject syncObj = null;
                    try
                    {
                        syncObj = syncObjects[i];
                        if (syncObj != null)
                        {
                            syncObj.SyncEnd += SyncObject_SyncEnd;
                            _syncObjectsList.Add(syncObj);
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Debug($"SyncObjects の購読中に例外が発生しました: {ex.Message}");
            }
            finally
            {
                MailAddressHelper.ReleaseCom(syncObjects);
                MailAddressHelper.ReleaseCom(session);
            }
        }

        private void InboxItems_ItemAdd(object item)
        {
            HandleItemAdd(item, isSentFolder: false);
        }

        private void SentItems_ItemAdd(object item)
        {
            HandleItemAdd(item, isSentFolder: true);
        }

        private void HandleItemAdd(object item, bool isSentFolder)
        {
            try
            {
                if (!SettingsManager.Current.EnableIncomingAndSentAutoProcess)
                {
                    return;
                }

                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    return;
                }

                IncomingMailProcessor.ProcessStoredMailItem(mail, isSentFolderItem: isSentFolder);
            }
            catch (Exception ex)
            {
                FileLogger.Error($"[ItemAdd] 新着アイテム処理中にエラーが発生しました (SentFolder={isSentFolder})。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(item);
            }
        }

        private void SyncObject_SyncEnd()
        {
            // 送受信完了時に差分スキャンをスケジュール
            if (SettingsManager.Current.EnableIncomingAndSentAutoProcess)
            {
                FileLogger.Debug("[SyncEnd] 送受信完了を検知したため差分スキャンを実行します。");
                EnqueueRecentUnprocessedItems();
            }
        }

        private void OnSettingsChanged(object sender, EventArgs e)
        {
            try
            {
                AddInSettings settings = SettingsManager.Current;
                if (_periodicTimer != null)
                {
                    int intervalMs = Math.Max(1, settings.PeriodicScanIntervalMinutes) * 60 * 1000;
                    _periodicTimer.Interval = intervalMs;
                }
            }
            catch
            {
            }
        }

        private void PeriodicTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                AddInSettings settings = SettingsManager.Current;
                int nextIntervalMs = Math.Max(1, settings.PeriodicScanIntervalMinutes) * 60 * 1000;
                if (_periodicTimer.Interval != nextIntervalMs)
                {
                    _periodicTimer.Interval = nextIntervalMs;
                }

                if (!settings.EnableIncomingAndSentAutoProcess)
                {
                    return;
                }

                EnqueueRecentUnprocessedItems();
            }
            catch (Exception ex)
            {
                FileLogger.Error("[定期スキャン] スキャン対象収集時にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// 直近 N 日間の受信トレイ・送信済みアイテムから未処理メールを抽出し、非同期チャンクキューへ追加します。
        /// </summary>
        public void EnqueueRecentUnprocessedItems()
        {
            if (_isCollectingCatchUp)
            {
                return;
            }

            _isCollectingCatchUp = true;
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;

            try
            {
                AddInSettings settings = SettingsManager.Current;
                DateTime cutoff = DateTime.Now.AddDays(-Math.Max(1, settings.AutoScanLookbackDays));

                session = _application.Session;
                stores = session.Stores;
                int storeCount = stores.Count;

                for (int i = 1; i <= storeCount; i++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        store = stores[i];
                        CollectFromDefaultFolder(store, Outlook.OlDefaultFolders.olFolderInbox, isSentFolder: false, cutoff);
                        CollectFromDefaultFolder(store, Outlook.OlDefaultFolders.olFolderSentMail, isSentFolder: true, cutoff);
                    }
                    catch
                    {
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(store);
                    }
                }

                if (_scanQueue.Count > 0 && !_chunkWorkerTimer.Enabled)
                {
                    FileLogger.Debug($"[差分スキャン] キューに {_scanQueue.Count} 件の候補を追加しました。バックグラウンド処理を開始します。");
                    _chunkWorkerTimer.Start();
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("[差分スキャン] 候補アイテム収集処理でエラーが発生しました。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(stores);
                MailAddressHelper.ReleaseCom(session);
                _isCollectingCatchUp = false;
            }
        }

        private void CollectFromDefaultFolder(
            Outlook.Store store,
            Outlook.OlDefaultFolders folderType,
            bool isSentFolder,
            DateTime cutoff)
        {
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            Outlook.Items restricted = null;

            try
            {
                folder = store.GetDefaultFolder(folderType);
                if (folder == null)
                {
                    return;
                }

                string storeId = folder.StoreID;
                items = folder.Items;

                // Restrict フィルタで直近 N 日間のアイテムのみに絞り込む
                string dateStr = cutoff.ToString("g", CultureInfo.CurrentCulture);
                string filter = $"[ReceivedTime] >= '{dateStr}'";
                restricted = items.Restrict(filter);

                int count = restricted.Count;
                for (int i = 1; i <= count; i++)
                {
                    object rawItem = null;
                    Outlook.PropertyAccessor pa = null;
                    try
                    {
                        rawItem = restricted[i];
                        Outlook.MailItem mail = rawItem as Outlook.MailItem;
                        if (mail == null)
                        {
                            continue;
                        }

                        pa = mail.PropertyAccessor;
                        string processed = MailAddressHelper.TryGetStringProperty(pa, MapiConstants.CUSTOM_PROCESSED_PROP);
                        if (processed == "1")
                        {
                            continue;
                        }

                        string entryId = mail.EntryID;
                        if (!string.IsNullOrEmpty(entryId) && _queuedEntryIds.Add(entryId))
                        {
                            _scanQueue.Enqueue(new PendingScanItem
                            {
                                EntryId = entryId,
                                StoreId = storeId,
                                IsSentFolder = isSentFolder
                            });
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(pa);
                        MailAddressHelper.ReleaseCom(rawItem);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                MailAddressHelper.ReleaseCom(restricted);
                MailAddressHelper.ReleaseCom(items);
                MailAddressHelper.ReleaseCom(folder);
            }
        }

        /// <summary>
        /// UI スレッドを固めないよう、1回の Tick につき最大 ItemsPerTick 件だけ処理します。
        /// </summary>
        private void ChunkWorkerTimer_Tick(object sender, EventArgs e)
        {
            if (!SettingsManager.Current.EnableIncomingAndSentAutoProcess)
            {
                _scanQueue.Clear();
                _queuedEntryIds.Clear();
                _chunkWorkerTimer.Stop();
                return;
            }

            if (_scanQueue.Count == 0)
            {
                _chunkWorkerTimer.Stop();
                return;
            }

            Outlook.NameSpace session = null;
            try
            {
                session = _application.Session;
                int processedThisTick = 0;

                while (_scanQueue.Count > 0 && processedThisTick < ItemsPerTick)
                {
                    PendingScanItem work = _scanQueue.Dequeue();
                    _queuedEntryIds.Remove(work.EntryId);

                    object rawItem = null;
                    try
                    {
                        rawItem = session.GetItemFromID(work.EntryId, work.StoreId);
                        Outlook.MailItem mail = rawItem as Outlook.MailItem;
                        if (mail != null)
                        {
                            IncomingMailProcessor.ProcessStoredMailItem(mail, isSentFolderItem: work.IsSentFolder);
                        }
                    }
                    catch
                    {
                        // 処理前に削除・移動されたアイテムは無視
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(rawItem);
                    }

                    processedThisTick++;
                }

                if (_scanQueue.Count == 0)
                {
                    _chunkWorkerTimer.Stop();
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("[差分スキャン] チャンク処理中にエラーが発生しました。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(session);
            }
        }

        public void Dispose()
        {
            SettingsManager.SettingsChanged -= OnSettingsChanged;

            if (_periodicTimer != null)
            {
                _periodicTimer.Stop();
                _periodicTimer.Tick -= PeriodicTimer_Tick;
                _periodicTimer.Dispose();
                _periodicTimer = null;
            }

            if (_chunkWorkerTimer != null)
            {
                _chunkWorkerTimer.Stop();
                _chunkWorkerTimer.Tick -= ChunkWorkerTimer_Tick;
                _chunkWorkerTimer.Dispose();
                _chunkWorkerTimer = null;
            }

            foreach (Outlook.Items items in _inboxItemsList)
            {
                try
                {
                    items.ItemAdd -= InboxItems_ItemAdd;
                }
                catch
                {
                }
                MailAddressHelper.ReleaseCom(items);
            }
            _inboxItemsList.Clear();

            foreach (Outlook.Items items in _sentItemsList)
            {
                try
                {
                    items.ItemAdd -= SentItems_ItemAdd;
                }
                catch
                {
                }
                MailAddressHelper.ReleaseCom(items);
            }
            _sentItemsList.Clear();

            foreach (Outlook.SyncObject syncObj in _syncObjectsList)
            {
                try
                {
                    syncObj.SyncEnd -= SyncObject_SyncEnd;
                }
                catch
                {
                }
                MailAddressHelper.ReleaseCom(syncObj);
            }
            _syncObjectsList.Clear();
        }
    }
}
