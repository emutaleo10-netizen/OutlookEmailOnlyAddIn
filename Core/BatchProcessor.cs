using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookEmailOnlyAddIn.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// 一括処理の対象フォルダ情報
    /// </summary>
    public sealed class BatchFolderTarget
    {
        public string FolderPath { get; set; }
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public bool TreatAsSentFolder { get; set; }
    }

    /// <summary>
    /// 一括処理の進捗レポートデータ
    /// </summary>
    public sealed class BatchProgressInfo
    {
        public string CurrentFolderPath { get; set; }
        public int ProcessedCount { get; set; }
        public int TotalCount { get; set; }
        public int ModifiedCount { get; set; }
        public int AlreadyProcessedCount { get; set; }
        public int UnchangedCount { get; set; }
        public int SkippedOrErrorCount { get; set; }
    }

    /// <summary>
    /// フォルダ単位の一括処理を UI スレッド（STA）上で安全かつ非同期チャンク実行するクラス。
    /// COM オブジェクトを1件ごとに解放し、一定件数ごとに Task.Delay で UI メッセージポンプへ制御を返すことで、
    /// Outlook のフリーズや PST 破損を防ぎつつ、いつでも中断・再開できるようにします。
    /// </summary>
    public sealed class BatchProcessor
    {
        private readonly Outlook.Application _application;

        // UI の応答性を維持するため、10件処理するごとにメッセージポンプへ制御を戻す
        private const int ChunkSize = 10;

        public BatchProcessor(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// 指定されたフォルダ一覧に対して一括書き換え処理を実行します（UI スレッドから呼び出してください）。
        /// </summary>
        public async Task<BatchProgressInfo> ExecuteAsync(
            IList<BatchFolderTarget> targets,
            bool forceReprocess,
            IProgress<BatchProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var stats = new BatchProgressInfo();
            if (targets == null || targets.Count == 0)
            {
                return stats;
            }

            FileLogger.Info($"[一括処理開始] 対象フォルダ数: {targets.Count}, 処理済み再実行: {forceReprocess}");

            // 1. まず対象フォルダ内の全アイテム EntryID リストを収集し、全体の母数を確定する
            var workList = new List<WorkEntry>();
            Outlook.NameSpace session = null;

            try
            {
                session = _application.Session;

                foreach (BatchFolderTarget target in targets)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    stats.CurrentFolderPath = target.FolderPath;
                    progress?.Report(CloneStats(stats));
                    await Task.Delay(10).ConfigureAwait(true);

                    CollectFolderEntryIds(session, target, workList, cancellationToken);
                }

                stats.TotalCount = workList.Count;
                progress?.Report(CloneStats(stats));

                FileLogger.Info($"[一括処理] 対象メール総数: {stats.TotalCount} 件の収集が完了しました。");

                // 2. チャンク単位で1件ずつ処理を実行する
                for (int i = 0; i < workList.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    WorkEntry entry = workList[i];
                    stats.CurrentFolderPath = entry.FolderPath;

                    object rawItem = null;
                    try
                    {
                        rawItem = session.GetItemFromID(entry.ItemEntryId, entry.StoreId);
                        Outlook.MailItem mail = rawItem as Outlook.MailItem;
                        if (mail != null)
                        {
                            MailProcessResult result = IncomingMailProcessor.ProcessStoredMailItem(
                                mail,
                                isSentFolderItem: entry.TreatAsSentFolder,
                                forceReprocess: forceReprocess);

                            switch (result)
                            {
                                case MailProcessResult.Modified:
                                    stats.ModifiedCount++;
                                    break;
                                case MailProcessResult.AlreadyProcessed:
                                    stats.AlreadyProcessedCount++;
                                    break;
                                case MailProcessResult.UnchangedMarked:
                                    stats.UnchangedCount++;
                                    break;
                                default:
                                    stats.SkippedOrErrorCount++;
                                    break;
                            }
                        }
                        else
                        {
                            stats.SkippedOrErrorCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        stats.SkippedOrErrorCount++;
                        FileLogger.Error($"[一括処理] メール処理中にエラーが発生しました (EntryID={entry.ItemEntryId})", ex);
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(rawItem);
                    }

                    stats.ProcessedCount = i + 1;

                    // ChunkSize 件ごと、または最終件で進捗更新と UI 応答のための Yield を行う
                    if ((i + 1) % ChunkSize == 0 || i == workList.Count - 1)
                    {
                        progress?.Report(CloneStats(stats));
                        await Task.Delay(15).ConfigureAwait(true);
                    }
                }

                FileLogger.Info(
                    $"[一括処理完了] 合計: {stats.ProcessedCount}/{stats.TotalCount} 件 " +
                    $"(書き換え: {stats.ModifiedCount}, 処理済みスキップ: {stats.AlreadyProcessedCount}, " +
                    $"変更なし: {stats.UnchangedCount}, 対象外/エラー: {stats.SkippedOrErrorCount})");
            }
            catch (OperationCanceledException)
            {
                FileLogger.Info(
                    $"[一括処理中断] ユーザーによりキャンセルされました。処理済み: {stats.ProcessedCount}/{stats.TotalCount} 件 " +
                    $"(書き換え: {stats.ModifiedCount} 件)。次回実行時は未処理分から再開可能です。");
                throw;
            }
            finally
            {
                MailAddressHelper.ReleaseCom(session);
            }

            return stats;
        }

        private sealed class WorkEntry
        {
            public string FolderPath { get; set; }
            public string ItemEntryId { get; set; }
            public string StoreId { get; set; }
            public bool TreatAsSentFolder { get; set; }
        }

        private static void CollectFolderEntryIds(
            Outlook.NameSpace session,
            BatchFolderTarget target,
            List<WorkEntry> outputList,
            CancellationToken cancellationToken)
        {
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;

            try
            {
                folder = session.GetFolderFromID(target.EntryId, target.StoreId);
                if (folder == null)
                {
                    return;
                }

                items = folder.Items;
                int count = items.Count;

                for (int i = 1; i <= count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    object rawItem = null;
                    try
                    {
                        rawItem = items[i];
                        Outlook.MailItem mail = rawItem as Outlook.MailItem;
                        if (mail != null)
                        {
                            string entryId = mail.EntryID;
                            if (!string.IsNullOrEmpty(entryId))
                            {
                                outputList.Add(new WorkEntry
                                {
                                    FolderPath = target.FolderPath,
                                    ItemEntryId = entryId,
                                    StoreId = target.StoreId,
                                    TreatAsSentFolder = target.TreatAsSentFolder
                                });
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(rawItem);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                FileLogger.Error($"[一括処理] フォルダ '{target.FolderPath}' のアイテム一覧取得中にエラーが発生しました。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(items);
                MailAddressHelper.ReleaseCom(folder);
            }
        }

        private static BatchProgressInfo CloneStats(BatchProgressInfo src)
        {
            return new BatchProgressInfo
            {
                CurrentFolderPath = src.CurrentFolderPath,
                ProcessedCount = src.ProcessedCount,
                TotalCount = src.TotalCount,
                ModifiedCount = src.ModifiedCount,
                AlreadyProcessedCount = src.AlreadyProcessedCount,
                UnchangedCount = src.UnchangedCount,
                SkippedOrErrorCount = src.SkippedOrErrorCount
            };
        }
    }
}
