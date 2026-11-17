using System;
using System.Collections.Generic;
using OutlookEmailOnlyAddIn.Configuration;
using OutlookEmailOnlyAddIn.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// メール1件の書き換え結果を表す列挙型
    /// </summary>
    public enum MailProcessResult
    {
        /// <summary>
        /// 表示名をメールアドレスに書き換えて保存した
        /// </summary>
        Modified = 0,

        /// <summary>
        /// 既に処理済みのためスキップした
        /// </summary>
        AlreadyProcessed = 1,

        /// <summary>
        /// 除外ドメインまたは書き換え不要のため変更なしで処理済みマークのみ付与した
        /// </summary>
        UnchangedMarked = 2,

        /// <summary>
        /// 署名・暗号化メールや下書き等のためスキップした
        /// </summary>
        SkippedUnsupported = 3,

        /// <summary>
        /// エラー発生によりスキップした
        /// </summary>
        Failed = 4
    }

    /// <summary>
    /// 受信トレイおよび送信済みアイテム内の MailItem に対して、
    /// 一覧に表示される差出人名・宛先名をメールアドレスのみに書き換えるプロセッサクラス
    /// </summary>
    public static class IncomingMailProcessor
    {
        // 同一プロセス内で現在処理中の EntryID を保持し、Save() による再入（無限ループ）を防止する
        private static readonly HashSet<string> ProcessingEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly object EntryIdLock = new object();

        /// <summary>
        /// 指定された MailItem を検査し、必要に応じて差出人表示名または宛先表示名をメールアドレスに書き換えます。
        /// </summary>
        /// <param name="mail">対象の MailItem</param>
        /// <param name="isSentFolderItem">送信済みアイテムフォルダ内のメール（宛先書き換え主対象）かどうか</param>
        /// <param name="forceReprocess">既に処理済みマークが付いていても再チェックするか（一括処理オプション用）</param>
        public static MailProcessResult ProcessStoredMailItem(
            Outlook.MailItem mail,
            bool isSentFolderItem,
            bool forceReprocess = false)
        {
            if (mail == null)
            {
                return MailProcessResult.SkippedUnsupported;
            }

            string entryId = null;
            string subject = "(件名なし)";

            try
            {
                entryId = mail.EntryID;
                subject = mail.Subject ?? "(件名なし)";
            }
            catch
            {
            }

            // 未送信（作成中・下書き・送信トレイ滞留中）のメールは受信/保存後処理の対象外とする
            try
            {
                if (!mail.Sent)
                {
                    return MailProcessResult.SkippedUnsupported;
                }
            }
            catch
            {
                return MailProcessResult.SkippedUnsupported;
            }

            // 再入防止チェック
            if (!string.IsNullOrEmpty(entryId))
            {
                lock (EntryIdLock)
                {
                    if (!ProcessingEntryIds.Add(entryId))
                    {
                        return MailProcessResult.AlreadyProcessed;
                    }
                }
            }

            try
            {
                // S/MIME（電子署名・暗号化）メールは破損防止のためスキップ
                if (MailAddressHelper.IsSignedOrEncrypted(mail))
                {
                    FileLogger.Debug($"[保存メールスキップ] 署名/暗号化メールのため対象外: {subject}");
                    return MailProcessResult.SkippedUnsupported;
                }

                Outlook.PropertyAccessor pa = null;
                try
                {
                    pa = mail.PropertyAccessor;

                    // 処理済みマークの確認
                    if (!forceReprocess)
                    {
                        string processedFlag = MailAddressHelper.TryGetStringProperty(pa, MapiConstants.CUSTOM_PROCESSED_PROP);
                        if (processedFlag == "1")
                        {
                            return MailProcessResult.AlreadyProcessed;
                        }
                    }

                    AddInSettings settings = SettingsManager.Current;
                    bool modified = false;

                    if (isSentFolderItem)
                    {
                        // 送信済みアイテム: 一覧の「宛先」列（PR_DISPLAY_TO / PR_DISPLAY_CC）をメールアドレス化するため
                        // Recipients 内の各宛先の表示名を書き換える
                        if (RewriteRecipientsDisplayNames(mail, settings))
                        {
                            modified = true;
                        }
                    }
                    else
                    {
                        // 受信メール: 一覧の「差出人」列（PR_SENT_REPRESENTING_NAME / PR_SENDER_NAME）をメールアドレス化する
                        if (RewriteSenderDisplayName(mail, pa, settings))
                        {
                            modified = true;
                        }

                        // オプション: 受信メールの宛先（To/Cc）の表示名も書き換える設定が有効な場合
                        if (settings.RewriteRecipientsOnIncomingMail)
                        {
                            if (RewriteRecipientsDisplayNames(mail, settings))
                            {
                                modified = true;
                            }
                        }
                    }

                    // 再スキャン時の重複処理を防ぐため処理済みマークを記録して保存
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.CUSTOM_PROCESSED_PROP, "1");
                    mail.Save();

                    if (modified)
                    {
                        FileLogger.Info($"[表示名書き換え完了] {(isSentFolderItem ? "送信済み宛先" : "受信差出人")}を更新しました。件名: {subject}");
                        return MailProcessResult.Modified;
                    }
                    else
                    {
                        FileLogger.Debug($"[表示名変更なし] 除外ドメインまたは変更不要のため処理済みマークのみ記録。件名: {subject}");
                        return MailProcessResult.UnchangedMarked;
                    }
                }
                finally
                {
                    MailAddressHelper.ReleaseCom(pa);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error($"[保存メール処理エラー] メールの表示名書き換え中にエラーが発生しました。件名: {subject}", ex);
                return MailProcessResult.Failed;
            }
            finally
            {
                if (!string.IsNullOrEmpty(entryId))
                {
                    lock (EntryIdLock)
                    {
                        ProcessingEntryIds.Remove(entryId);
                    }
                }
            }
        }

        /// <summary>
        /// 受信メールの差出人（From）の表示名を SMTP アドレスに書き換えます。
        /// </summary>
        private static bool RewriteSenderDisplayName(
            Outlook.MailItem mail,
            Outlook.PropertyAccessor mailPa,
            AddInSettings settings)
        {
            string smtp = MailAddressHelper.GetSenderSmtpAddress(mail);
            if (!MailAddressHelper.IsValidSmtpAddress(smtp))
            {
                FileLogger.Debug("[差出人判定] 有効な SMTP アドレスが取得できないため元の表示名を維持します。");
                return false;
            }

            if (DomainFilter.IsExcludedDomain(smtp, settings))
            {
                FileLogger.Debug($"[差出人判定] 除外ドメインのため元の表示名を維持します: {smtp}");
                return false;
            }

            string currentSenderName = MailAddressHelper.TryGetStringProperty(mailPa, MapiConstants.PR_SENDER_NAME);
            string currentReprName = MailAddressHelper.TryGetStringProperty(mailPa, MapiConstants.PR_SENT_REPRESENTING_NAME);

            bool needsUpdate =
                !string.Equals(currentSenderName, smtp, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(currentReprName, smtp, StringComparison.OrdinalIgnoreCase);

            if (!needsUpdate)
            {
                return false;
            }

            // 元の表示名を退避プロパティに保存（未保存の場合のみ）
            string existingBackup = MailAddressHelper.TryGetStringProperty(mailPa, MapiConstants.CUSTOM_ORIGINAL_SENDER_NAME_PROP);
            if (string.IsNullOrEmpty(existingBackup))
            {
                string backupName = !string.IsNullOrEmpty(currentReprName) ? currentReprName : (currentSenderName ?? string.Empty);
                if (!string.IsNullOrEmpty(backupName))
                {
                    MailAddressHelper.TrySetProperty(mailPa, MapiConstants.CUSTOM_ORIGINAL_SENDER_NAME_PROP, backupName);
                }
            }

            // 一覧表示および閲覧画面で参照される差出人名プロパティを SMTP アドレスで上書き
            bool set1 = MailAddressHelper.TrySetProperty(mailPa, MapiConstants.PR_SENDER_NAME, smtp);
            bool set2 = MailAddressHelper.TrySetProperty(mailPa, MapiConstants.PR_SENT_REPRESENTING_NAME, smtp);
            MailAddressHelper.TrySetProperty(mailPa, MapiConstants.PR_ORIGINAL_SENDER_NAME, smtp);
            MailAddressHelper.TrySetProperty(mailPa, MapiConstants.PR_ORIGINAL_SENT_REPRESENTING_NAME, smtp);

            return set1 || set2;
        }

        /// <summary>
        /// MailItem.Recipients 内の各宛先の PR_DISPLAY_NAME を SMTP アドレスに書き換えます。
        /// 送信済みアイテムでこれを実行して mail.Save() すると、PST プロバイダが PR_DISPLAY_TO / PR_DISPLAY_CC を
        /// メールアドレスのみで再計算し、一覧の「宛先」列がメールアドレス表示になります。
        /// </summary>
        private static bool RewriteRecipientsDisplayNames(Outlook.MailItem mail, AddInSettings settings)
        {
            bool anyModified = false;
            Outlook.Recipients recipients = null;

            try
            {
                recipients = mail.Recipients;
                if (recipients == null || recipients.Count == 0)
                {
                    return false;
                }

                int count = recipients.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Recipient r = null;
                    Outlook.PropertyAccessor rPa = null;
                    try
                    {
                        r = recipients[i];
                        int type = r.Type;

                        if (type == (int)Outlook.OlMailRecipientType.olBCC && !settings.IncludeBccInOutgoing)
                        {
                            continue;
                        }

                        string smtp = MailAddressHelper.GetRecipientSmtpAddress(r);
                        if (!MailAddressHelper.IsValidSmtpAddress(smtp))
                        {
                            continue;
                        }

                        if (DomainFilter.IsExcludedDomain(smtp, settings))
                        {
                            continue;
                        }

                        rPa = r.PropertyAccessor;
                        string currentDisp = MailAddressHelper.TryGetStringProperty(rPa, MapiConstants.PR_DISPLAY_NAME);

                        // 連絡先エントリへの紐づけ（EntryID）を One-Off EntryID に差し替え、表示名をメールアドレス化する
                        byte[] oneOffEntryId = MailAddressHelper.CreateOneOffEntryId(smtp);
                        if (oneOffEntryId != null)
                        {
                            MailAddressHelper.TrySetProperty(rPa, MapiConstants.PR_ENTRYID, oneOffEntryId);
                        }

                        if (!string.Equals(currentDisp, smtp, StringComparison.OrdinalIgnoreCase))
                        {
                            bool ok1 = MailAddressHelper.TrySetProperty(rPa, MapiConstants.PR_DISPLAY_NAME, smtp);
                            bool ok2 = MailAddressHelper.TrySetProperty(rPa, MapiConstants.PR_TRANSMITABLE_DISPLAY_NAME, smtp);
                            if (ok1 || ok2)
                            {
                                anyModified = true;
                            }
                        }
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(rPa);
                        MailAddressHelper.ReleaseCom(r);
                    }
                }
            }
            finally
            {
                MailAddressHelper.ReleaseCom(recipients);
            }

            return anyModified;
        }
    }
}
