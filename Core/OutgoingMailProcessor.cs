using System;
using System.Collections.Generic;
using OutlookEmailOnlyAddIn.Configuration;
using OutlookEmailOnlyAddIn.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// 送信時（Application.ItemSend）にメールの宛先（To / Cc / 設定により Bcc）から
    /// 表示名を除去し、メールアドレスのみに再構築するプロセッサクラス
    /// </summary>
    public static class OutgoingMailProcessor
    {
        /// <summary>
        /// 再構築用の宛先情報モデル
        /// </summary>
        private sealed class RecipientWorkItem
        {
            public int RecipientType { get; set; }
            public string OriginalName { get; set; }
            public string SmtpAddress { get; set; }
            public bool ShouldStripDisplayName { get; set; }
        }

        /// <summary>
        /// ItemSend イベントで渡された送信アイテムを処理します。
        /// 例外が発生した場合でもメール送信自体は妨げない（フェイルセーフ）設計としています。
        /// </summary>
        /// <param name="item">送信対象のアイテム</param>
        public static void ProcessOutgoingItem(object item)
        {
            Outlook.MailItem mail = item as Outlook.MailItem;
            if (mail == null)
            {
                return;
            }

            AddInSettings settings = SettingsManager.Current;
            if (!settings.EnableOutgoingProcess)
            {
                FileLogger.Debug("送信処理は無効化されているためスキップしました。");
                return;
            }

            string subject = SafeGetSubject(mail);

            try
            {
                // S/MIME（署名・暗号化）メールの場合は安全のため対象外とする
                if (MailAddressHelper.IsSignedOrEncrypted(mail))
                {
                    FileLogger.Info($"[送信スキップ] 署名または暗号化メールのため対象外としました。件名: {subject}");
                    return;
                }

                Outlook.Recipients recipients = null;
                try
                {
                    recipients = mail.Recipients;
                    if (recipients == null || recipients.Count == 0)
                    {
                        return;
                    }

                    // 1. まず連絡先グループ（配布リスト）の展開処理を行う（設定ONの場合）
                    if (settings.ExpandDistributionListOnSend)
                    {
                        ExpandDistributionListsIfNeeded(recipients, settings, subject);
                    }

                    // 2. 全宛先を走査し、表示名除去が必要かどうかを判定する
                    var workItems = new List<RecipientWorkItem>();
                    bool anyNeedsRewrite = false;

                    int count = recipients.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Recipient r = null;
                        try
                        {
                            r = recipients[i];
                            RecipientWorkItem info = AnalyzeRecipient(r, settings);
                            workItems.Add(info);

                            if (info.ShouldStripDisplayName)
                            {
                                anyNeedsRewrite = true;
                            }
                        }
                        finally
                        {
                            MailAddressHelper.ReleaseCom(r);
                        }
                    }

                    if (!anyNeedsRewrite)
                    {
                        FileLogger.Info($"[送信] 書き換え対象の宛先はありませんでした。件名: {subject}");
                        return;
                    }

                    // 3. 書き換え対象が含まれる場合、対象の Recipient を削除して純粋なメールアドレスで再構築する
                    //    ※既存の連絡先解決済み Recipient の PR_DISPLAY_NAME だけを変更しても
                    //      内部の EntryID（連絡先リンク）が残っていると送信時に元の表示名が復活するため、
                    //      対象宛先を Remove & Add し、MAPI プロパティを明示設定する方式と併用する。
                    int rewrittenCount = RebuildTargetRecipients(recipients, workItems, subject);

                    // 4. 変更を確定し、カスタム処理済みマークを付与
                    MarkMailProcessed(mail);
                    mail.Save();

                    FileLogger.Info($"[送信完了] 宛先表示名を除去しました（書き換え件数: {rewrittenCount}件）。件名: {subject}");
                }
                finally
                {
                    MailAddressHelper.ReleaseCom(recipients);
                }
            }
            catch (Exception ex)
            {
                // フェイルセーフ: 例外発生時も Cancel は変更せず、元の表示名のまま送信を継続させる
                FileLogger.Error($"[送信エラー] 宛先表示名の除去処理中に例外が発生しました。元の状態で送信を継続します。件名: {subject}", ex);
            }
        }

        /// <summary>
        /// 宛先1件を解析し、表示名除去の対象かを判定します。
        /// </summary>
        private static RecipientWorkItem AnalyzeRecipient(Outlook.Recipient recipient, AddInSettings settings)
        {
            int type = (int)Outlook.OlMailRecipientType.olTo;
            string origName = string.Empty;

            try
            {
                type = recipient.Type;
                origName = recipient.Name ?? string.Empty;
            }
            catch
            {
            }

            string smtp = MailAddressHelper.GetRecipientSmtpAddress(recipient);

            var item = new RecipientWorkItem
            {
                RecipientType = type,
                OriginalName = origName,
                SmtpAddress = smtp,
                ShouldStripDisplayName = false
            };

            // Bcc の対象判定
            if (type == (int)Outlook.OlMailRecipientType.olBCC && !settings.IncludeBccInOutgoing)
            {
                FileLogger.Debug($"[送信判定] Bcc は設定により対象外: {origName} ({smtp})");
                return item;
            }

            // To / Cc / Bcc 以外（Originator 等）は対象外
            if (type != (int)Outlook.OlMailRecipientType.olTo &&
                type != (int)Outlook.OlMailRecipientType.olCC &&
                type != (int)Outlook.OlMailRecipientType.olBCC)
            {
                return item;
            }

            // 有効な SMTP アドレスが取得できない場合は元の表示名を維持
            if (!MailAddressHelper.IsValidSmtpAddress(smtp))
            {
                FileLogger.Info($"[送信判定] SMTPアドレスが取得できないか形式不正のため維持: {origName}");
                return item;
            }

            // ローカルドメイン（@kinki-sharyo.local 等）は除外
            if (DomainFilter.IsExcludedDomain(smtp, settings))
            {
                FileLogger.Debug($"[送信判定] 除外ドメインのため表示名を維持: {origName} <{smtp}>");
                return item;
            }

            // 表示名がすでにメールアドレスと完全一致している場合でも、
            // 連絡先 EntryID が紐づいていると SMTP ヘッダー構築時に連絡先名が出ることがあるため、
            // 除外ドメイン以外の有効な SMTP 宛先はすべてメールアドレスのみへ正規化する。
            item.ShouldStripDisplayName = true;
            return item;
        }

        /// <summary>
        /// 書き換え対象の Recipient を削除し、純粋なメールアドレスのみで再登録・プロパティ設定を行います。
        /// </summary>
        private static int RebuildTargetRecipients(
            Outlook.Recipients recipients,
            List<RecipientWorkItem> workItems,
            string subject)
        {
            int rewrittenCount = 0;

            // 削除時にインデックスがずれないよう逆順で Remove する
            for (int i = workItems.Count - 1; i >= 0; i--)
            {
                if (workItems[i].ShouldStripDisplayName)
                {
                    try
                    {
                        recipients.Remove(i + 1);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Error($"[送信] 宛先の削除に失敗しました (Index={i + 1}, Address={workItems[i].SmtpAddress})", ex);
                        workItems[i].ShouldStripDisplayName = false;
                    }
                }
            }

            // 削除した宛先を元の順序どおり純粋な SMTP アドレスで再追加する
            foreach (RecipientWorkItem item in workItems)
            {
                if (!item.ShouldStripDisplayName)
                {
                    continue;
                }

                Outlook.Recipient newRecip = null;
                Outlook.PropertyAccessor pa = null;
                try
                {
                    newRecip = recipients.Add(item.SmtpAddress);
                    newRecip.Type = item.RecipientType;

                    // Resolve() を呼ぶとローカル連絡先に一致した際に連絡先の EntryID と登録名（「A様」）が紐づくため、
                    // Resolve() 実行直後に MAPI One-Off EntryID（アドレス帳非参照フラグ OOP_DONT_LOOKUP 付き）を生成して
                    // PR_ENTRYID を差し替え、連絡先とのリンクを完全に切断したうえで表示名をメールアドレスに固定する。
                    newRecip.Resolve();

                    pa = newRecip.PropertyAccessor;
                    byte[] oneOffEntryId = MailAddressHelper.CreateOneOffEntryId(item.SmtpAddress);
                    if (oneOffEntryId != null)
                    {
                        MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_ENTRYID, oneOffEntryId);
                    }

                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_ADDRTYPE, "SMTP");
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_EMAIL_ADDRESS, item.SmtpAddress);
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_SMTP_ADDRESS, item.SmtpAddress);
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_DISPLAY_NAME, item.SmtpAddress);
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_TRANSMITABLE_DISPLAY_NAME, item.SmtpAddress);
                    MailAddressHelper.TrySetProperty(pa, MapiConstants.PR_SEND_RICH_INFO, false);

                    rewrittenCount++;
                    FileLogger.Debug($"[送信書き換え] '{item.OriginalName}' -> '{item.SmtpAddress}' (Type={item.RecipientType})");
                }
                catch (Exception ex)
                {
                    FileLogger.Error($"[送信] 宛先の再追加・プロパティ設定に失敗しました ({item.SmtpAddress})。件名: {subject}", ex);
                }
                finally
                {
                    MailAddressHelper.ReleaseCom(pa);
                    MailAddressHelper.ReleaseCom(newRecip);
                }
            }

            return rewrittenCount;
        }

        /// <summary>
        /// 宛先に含まれる連絡先グループ（配布リスト）を個別のメンバーアドレスに展開します。
        /// </summary>
        private static void ExpandDistributionListsIfNeeded(
            Outlook.Recipients recipients,
            AddInSettings settings,
            string subject)
        {
            // ネストした連絡先グループにも対応するため最大3階層まで展開を試みる
            for (int depth = 0; depth < 3; depth++)
            {
                bool expandedAny = false;
                int count = recipients.Count;

                for (int i = count; i >= 1; i--)
                {
                    Outlook.Recipient r = null;
                    Outlook.AddressEntry entry = null;
                    Outlook.AddressEntries members = null;

                    try
                    {
                        r = recipients[i];
                        int recipType = r.Type;

                        if (recipType == (int)Outlook.OlMailRecipientType.olBCC && !settings.IncludeBccInOutgoing)
                        {
                            continue;
                        }

                        entry = r.AddressEntry;
                        if (entry == null)
                        {
                            continue;
                        }

                        Outlook.OlDisplayType displayType = entry.DisplayType;
                        bool isDistList =
                            displayType == Outlook.OlDisplayType.olDistList ||
                            displayType == Outlook.OlDisplayType.olPrivateDistList;

                        if (!isDistList)
                        {
                            continue;
                        }

                        members = entry.Members;
                        if (members == null || members.Count == 0)
                        {
                            continue;
                        }

                        string groupName = r.Name ?? "(連絡先グループ)";
                        int memberCount = members.Count;
                        var memberAddresses = new List<string>();

                        for (int m = 1; m <= memberCount; m++)
                        {
                            Outlook.AddressEntry memberEntry = null;
                            try
                            {
                                memberEntry = members[m];
                                if (memberEntry != null)
                                {
                                    string smtp = MailAddressHelper.GetAddressEntrySmtpAddress(memberEntry);
                                    if (!string.IsNullOrEmpty(smtp))
                                    {
                                        memberAddresses.Add(smtp);
                                    }
                                    else if (!string.IsNullOrEmpty(memberEntry.Name))
                                    {
                                        memberAddresses.Add(memberEntry.Name);
                                    }
                                }
                            }
                            finally
                            {
                                MailAddressHelper.ReleaseCom(memberEntry);
                            }
                        }

                        if (memberAddresses.Count > 0)
                        {
                            recipients.Remove(i);
                            foreach (string addr in memberAddresses)
                            {
                                Outlook.Recipient added = null;
                                try
                                {
                                    added = recipients.Add(addr);
                                    added.Type = recipType;
                                    added.Resolve();
                                }
                                finally
                                {
                                    MailAddressHelper.ReleaseCom(added);
                                }
                            }

                            expandedAny = true;
                            FileLogger.Info($"[送信] 連絡先グループ '{groupName}' を {memberAddresses.Count} 件のメンバーに展開しました。件名: {subject}");
                        }
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Error($"[送信] 連絡先グループの展開中にエラーが発生しました。件名: {subject}", ex);
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(members);
                        MailAddressHelper.ReleaseCom(entry);
                        MailAddressHelper.ReleaseCom(r);
                    }
                }

                if (!expandedAny)
                {
                    break;
                }
            }
        }

        private static void MarkMailProcessed(Outlook.MailItem mail)
        {
            Outlook.PropertyAccessor pa = null;
            try
            {
                pa = mail.PropertyAccessor;
                MailAddressHelper.TrySetProperty(pa, MapiConstants.CUSTOM_PROCESSED_PROP, "1");
            }
            catch
            {
            }
            finally
            {
                MailAddressHelper.ReleaseCom(pa);
            }
        }

        private static string SafeGetSubject(Outlook.MailItem mail)
        {
            try
            {
                return mail.Subject ?? "(件名なし)";
            }
            catch
            {
                return "(件名取得不可)";
            }
        }
    }
}
