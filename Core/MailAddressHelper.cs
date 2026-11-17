using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// Outlook の MailItem / Recipient / AddressEntry から安全に SMTP アドレスを抽出・検証するヘルパークラス
    /// </summary>
    public static class MailAddressHelper
    {
        // 基本的な SMTP メールアドレス形式の正規表現（ローカルドメイン .local 等も許容）
        private static readonly Regex SmtpRegex = new Regex(
            @"^[^\s@<>""']+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // From ヘッダーからメールアドレスを抽出する正規表現
        private static readonly Regex HeaderFromEmailRegex = new Regex(
            @"^From:\s*.*?(?:<([^<>@\s]+@[^<>@\s]+)>|([^\s<>@""']+@[A-Za-z0-9.-]+))",
            RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);

        /// <summary>
        /// 文字列が有効な SMTP アドレス形式かどうかを判定します。
        /// </summary>
        public static bool IsValidSmtpAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            string candidate = address.Trim();
            // Exchange DN 形式（/O=...）は除外
            if (candidate.StartsWith("/", StringComparison.Ordinal))
            {
                return false;
            }

            return SmtpRegex.IsMatch(candidate);
        }

        /// <summary>
        /// Recipient オブジェクトから SMTP メールアドレスを取得します。
        /// </summary>
        public static string GetRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            if (recipient == null)
            {
                return null;
            }

            // 1. PropertyAccessor から PR_SMTP_ADDRESS を取得
            Outlook.PropertyAccessor pa = null;
            try
            {
                pa = recipient.PropertyAccessor;
                string smtp = TryGetStringProperty(pa, MapiConstants.PR_SMTP_ADDRESS);
                if (IsValidSmtpAddress(smtp))
                {
                    return smtp.Trim();
                }

                string orgEmail = TryGetStringProperty(pa, MapiConstants.PR_ORG_EMAIL_ADDR);
                if (IsValidSmtpAddress(orgEmail))
                {
                    return orgEmail.Trim();
                }

                string emailProp = TryGetStringProperty(pa, MapiConstants.PR_EMAIL_ADDRESS);
                if (IsValidSmtpAddress(emailProp))
                {
                    return emailProp.Trim();
                }
            }
            catch
            {
                // PropertyAccessor 失敗時は次へフォールバック
            }
            finally
            {
                ReleaseCom(pa);
            }

            // 2. Recipient.Address プロパティを確認
            try
            {
                string addr = recipient.Address;
                if (IsValidSmtpAddress(addr))
                {
                    return addr.Trim();
                }
            }
            catch
            {
            }

            // 3. Recipient.AddressEntry から取得
            Outlook.AddressEntry addressEntry = null;
            try
            {
                addressEntry = recipient.AddressEntry;
                if (addressEntry != null)
                {
                    string entrySmtp = GetAddressEntrySmtpAddress(addressEntry);
                    if (IsValidSmtpAddress(entrySmtp))
                    {
                        return entrySmtp.Trim();
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseCom(addressEntry);
            }

            // 4. 表示名自体がメールアドレス形式の場合（未解決の手入力直後など）
            try
            {
                string name = recipient.Name;
                if (IsValidSmtpAddress(name))
                {
                    return name.Trim();
                }

                // "A様 <a@example.com>" のような文字列から <> 内を抽出
                string extracted = ExtractBracketEmail(name);
                if (IsValidSmtpAddress(extracted))
                {
                    return extracted;
                }
            }
            catch
            {
            }

            return null;
        }

        /// <summary>
        /// AddressEntry オブジェクトから SMTP メールアドレスを取得します。
        /// </summary>
        public static string GetAddressEntrySmtpAddress(Outlook.AddressEntry addressEntry)
        {
            if (addressEntry == null)
            {
                return null;
            }

            Outlook.PropertyAccessor pa = null;
            try
            {
                pa = addressEntry.PropertyAccessor;
                string smtp = TryGetStringProperty(pa, MapiConstants.PR_SMTP_ADDRESS);
                if (IsValidSmtpAddress(smtp))
                {
                    return smtp.Trim();
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseCom(pa);
            }

            try
            {
                if (IsValidSmtpAddress(addressEntry.Address))
                {
                    return addressEntry.Address.Trim();
                }
            }
            catch
            {
            }

            return null;
        }

        /// <summary>
        /// MailItem から差出人（Sender）の SMTP メールアドレスを取得します。
        /// </summary>
        public static string GetSenderSmtpAddress(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return null;
            }

            // 1. SenderEmailAddress が直接 SMTP 形式の場合（POP/SMTP 環境では通常ここで取得可能）
            try
            {
                string senderEmail = mail.SenderEmailAddress;
                if (IsValidSmtpAddress(senderEmail))
                {
                    return senderEmail.Trim();
                }
            }
            catch
            {
            }

            // 2. PropertyAccessor から各種差出人アドレスプロパティを確認
            Outlook.PropertyAccessor pa = null;
            try
            {
                pa = mail.PropertyAccessor;

                string smtp = TryGetStringProperty(pa, MapiConstants.PR_SMTP_ADDRESS);
                if (IsValidSmtpAddress(smtp))
                {
                    return smtp.Trim();
                }

                string senderAddr = TryGetStringProperty(pa, MapiConstants.PR_SENDER_EMAIL_ADDRESS);
                if (IsValidSmtpAddress(senderAddr))
                {
                    return senderAddr.Trim();
                }

                string reprAddr = TryGetStringProperty(pa, MapiConstants.PR_SENT_REPRESENTING_EMAIL_ADDRESS);
                if (IsValidSmtpAddress(reprAddr))
                {
                    return reprAddr.Trim();
                }

                // 3. インターネットメールヘッダー（PR_TRANSPORT_MESSAGE_HEADERS）の From: 行から抽出
                string headers = TryGetStringProperty(pa, MapiConstants.PR_TRANSPORT_MESSAGE_HEADERS);
                if (!string.IsNullOrEmpty(headers))
                {
                    Match m = HeaderFromEmailRegex.Match(headers);
                    if (m.Success)
                    {
                        string extracted = !string.IsNullOrEmpty(m.Groups[1].Value)
                            ? m.Groups[1].Value
                            : m.Groups[2].Value;
                        if (IsValidSmtpAddress(extracted))
                        {
                            return extracted.Trim();
                        }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseCom(pa);
            }

            return null;
        }

        /// <summary>
        /// S/MIME（電子署名・暗号化）メールかどうかを判定します。
        /// </summary>
        public static bool IsSignedOrEncrypted(Outlook.MailItem mail)
        {
            if (mail == null)
            {
                return false;
            }

            try
            {
                string msgClass = mail.MessageClass;
                if (!string.IsNullOrEmpty(msgClass) &&
                    msgClass.IndexOf("SMIME", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            catch
            {
            }

            Outlook.PropertyAccessor pa = null;
            try
            {
                pa = mail.PropertyAccessor;
                object flagsObj = pa.GetProperty(MapiConstants.PR_SECURITY_FLAGS);
                if (flagsObj != null)
                {
                    int flags = Convert.ToInt32(flagsObj);
                    // 下位ビット（0x1: 暗号化, 0x2: 署名付き）が立っている場合は対象外とする
                    if ((flags & 0x03) != 0)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // プロパティが存在しない場合は通常の平文メール
            }
            finally
            {
                ReleaseCom(pa);
            }

            return false;
        }

        /// <summary>
        /// PropertyAccessor から文字列プロパティを安全に取得します。
        /// </summary>
        public static string TryGetStringProperty(Outlook.PropertyAccessor pa, string schemaName)
        {
            if (pa == null)
            {
                return null;
            }

            try
            {
                object val = pa.GetProperty(schemaName);
                return val as string;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// PropertyAccessor へプロパティを安全に設定します。
        /// </summary>
        public static bool TrySetProperty(Outlook.PropertyAccessor pa, string schemaName, object value)
        {
            if (pa == null)
            {
                return false;
            }

            try
            {
                pa.SetProperty(schemaName, value);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// "Display Name <user@example.com>" のような文字列から <> 内のメールアドレスを抽出します。
        /// </summary>
        public static string ExtractBracketEmail(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            int start = text.LastIndexOf('<');
            int end = text.LastIndexOf('>');
            if (start >= 0 && end > start + 1)
            {
                return text.Substring(start + 1, end - start - 1).Trim();
            }

            return null;
        }

        /// <summary>
        /// アドレス帳（連絡先）の登録名へのリンクを完全に断ち切るための
        /// MAPI One-Off EntryID ([MS-OXCDATA] 2.2.5.1) バイナリ配列を生成します。
        /// 表示名とメールアドレスの両方に純粋な SMTP アドレスを設定します。
        /// </summary>
        public static byte[] CreateOneOffEntryId(string smtpAddress)
        {
            if (string.IsNullOrWhiteSpace(smtpAddress))
            {
                return null;
            }

            string cleanAddr = smtpAddress.Trim();

            using (var ms = new System.IO.MemoryStream())
            using (var bw = new System.IO.BinaryWriter(ms))
            {
                // 1. Flags (4 bytes): 0x00000000
                bw.Write((uint)0);

                // 2. ProviderUID (16 bytes): MUIDOOP {812B1FA4-BEA3-1019-9D6E-00DD010F5402}
                byte[] providerUid = new byte[]
                {
                    0x81, 0x2B, 0x1F, 0xA4, 0xBE, 0xA3, 0x10, 0x19,
                    0x9D, 0x6E, 0x00, 0xDD, 0x01, 0x0F, 0x54, 0x02
                };
                bw.Write(providerUid);

                // 3. Version (2 bytes): 0x0000
                bw.Write((ushort)0);

                // 4. Flags (2 bytes): MAPI_UNICODE (0x8000) | OOP_DONT_LOOKUP (0x0100) | MAPI_SEND_NO_RICH_INFO (0x0001) = 0x8101
                bw.Write((ushort)0x8101);

                // 5. DisplayName (UTF-16LE, null-terminated): メールアドレスそのものを設定
                byte[] dispBytes = System.Text.Encoding.Unicode.GetBytes(cleanAddr);
                bw.Write(dispBytes);
                bw.Write((ushort)0);

                // 6. AddressType (UTF-16LE, null-terminated): "SMTP"
                byte[] typeBytes = System.Text.Encoding.Unicode.GetBytes("SMTP");
                bw.Write(typeBytes);
                bw.Write((ushort)0);

                // 7. EmailAddress (UTF-16LE, null-terminated): メールアドレス
                byte[] addrBytes = System.Text.Encoding.Unicode.GetBytes(cleanAddr);
                bw.Write(addrBytes);
                bw.Write((ushort)0);

                return ms.ToArray();
            }
        }

        /// <summary>
        /// COM オブジェクトを安全に解放します。
        /// </summary>
        public static void ReleaseCom(object comObj)
        {
            if (comObj != null && Marshal.IsComObject(comObj))
            {
                try
                {
                    Marshal.ReleaseComObject(comObj);
                }
                catch
                {
                }
            }
        }
    }
}
