namespace OutlookEmailOnlyAddIn.Core
{
    /// <summary>
    /// Outlook PropertyAccessor で使用する MAPI プロパティの DASL スキーマ定義
    /// </summary>
    public static class MapiConstants
    {
        /// <summary>
        /// PR_ENTRYID (0x0FFF0102) : オブジェクトまたはアドレス帳エントリの EntryID (PT_BINARY)
        /// </summary>
        public const string PR_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102";

        /// <summary>
        /// PR_ADDRTYPE (0x3002001F) : アドレス種別 ("SMTP" 等)
        /// </summary>
        public const string PR_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x3002001F";

        /// <summary>
        /// PR_DISPLAY_NAME (0x3001001F) : Recipient の表示名
        /// </summary>
        public const string PR_DISPLAY_NAME = "http://schemas.microsoft.com/mapi/proptag/0x3001001F";

        /// <summary>
        /// PR_TRANSMITABLE_DISPLAY_NAME (0x3A20001F) : 送信時に転送される表示名
        /// </summary>
        public const string PR_TRANSMITABLE_DISPLAY_NAME = "http://schemas.microsoft.com/mapi/proptag/0x3A20001F";

        /// <summary>
        /// PR_EMAIL_ADDRESS (0x3003001F) : アドレス文字列
        /// </summary>
        public const string PR_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x3003001F";

        /// <summary>
        /// PR_SMTP_ADDRESS (0x39FE001F) : SMTP メールアドレス
        /// </summary>
        public const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

        /// <summary>
        /// PR_ORG_EMAIL_ADDR (0x403E001F) : 連絡先から解決された宛先の元メールアドレス等
        /// </summary>
        public const string PR_ORG_EMAIL_ADDR = "http://schemas.microsoft.com/mapi/proptag/0x403E001F";

        /// <summary>
        /// PR_SENDER_NAME (0x0C1A001F) : 差出人表示名
        /// </summary>
        public const string PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F";

        /// <summary>
        /// PR_SENDER_EMAIL_ADDRESS (0x0C1F001F) : 差出人メールアドレス
        /// </summary>
        public const string PR_SENDER_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F";

        /// <summary>
        /// PR_SENT_REPRESENTING_NAME (0x0042001F) : 代理人/表示上の差出人名（一覧の差出人列で主に参照される）
        /// </summary>
        public const string PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F";

        /// <summary>
        /// PR_SENT_REPRESENTING_EMAIL_ADDRESS (0x0065001F) : 代理人/表示上の差出人メールアドレス
        /// </summary>
        public const string PR_SENT_REPRESENTING_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";

        /// <summary>
        /// PR_ORIGINAL_SENDER_NAME (0x005A001F) : 元の差出人表示名
        /// </summary>
        public const string PR_ORIGINAL_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x005A001F";

        /// <summary>
        /// PR_ORIGINAL_SENT_REPRESENTING_NAME (0x004D001F) : 元の表示差出人名
        /// </summary>
        public const string PR_ORIGINAL_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x004D001F";

        /// <summary>
        /// PR_TRANSPORT_MESSAGE_HEADERS (0x007D001F) : インターネットメールヘッダー全文
        /// </summary>
        public const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";

        /// <summary>
        /// PR_SECURITY_FLAGS (0x6E010003) : 署名・暗号化フラグ（0以外はS/MIME等のセキュリティ設定あり）
        /// </summary>
        public const string PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";

        /// <summary>
        /// PR_SEND_RICH_INFO (0x3A40000B) : リッチテキスト形式（TNEF）で送信するかどうかのフラグ
        /// </summary>
        public const string PR_SEND_RICH_INFO = "http://schemas.microsoft.com/mapi/proptag/0x3A40000B";

        /// <summary>
        /// 当アドインによる処理済み目印（PS_PUBLIC_STRINGS 名前空間のカスタム文字列プロパティ）
        /// UserProperties.Add を使用しないことで TNEF(winmail.dat) 化やフォーム定義肥大化を防止する
        /// </summary>
        public const string CUSTOM_PROCESSED_PROP =
            "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/EmailOnlyProcessed/0x0000001F";

        /// <summary>
        /// 書き換え前の差出人表示名を退避するためのカスタム文字列プロパティ
        /// </summary>
        public const string CUSTOM_ORIGINAL_SENDER_NAME_PROP =
            "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/EmailOnlyOrigSender/0x0000001F";
    }
}
