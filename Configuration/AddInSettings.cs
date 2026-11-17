using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OutlookEmailOnlyAddIn.Configuration
{
    /// <summary>
    /// ログ出力の詳細度を表す列挙型
    /// </summary>
    public enum LogLevel
    {
        Error = 0,
        Info = 1,
        Debug = 2
    }

    /// <summary>
    /// アドインのユーザー設定を保持するデータクラス
    /// </summary>
    [DataContract]
    public class AddInSettings
    {
        /// <summary>
        /// 送信メールの宛先表示名除去を有効にするか（既定: true）
        /// </summary>
        [DataMember]
        public bool EnableOutgoingProcess { get; set; } = true;

        /// <summary>
        /// 送信時に Bcc も表示名除去の対象にするか（既定: true）
        /// </summary>
        [DataMember]
        public bool IncludeBccInOutgoing { get; set; } = true;

        /// <summary>
        /// 送信時に連絡先グループ（配布リスト）を個別アドレスに展開して処理するか（既定: true）
        /// </summary>
        [DataMember]
        public bool ExpandDistributionListOnSend { get; set; } = true;

        /// <summary>
        /// 受信トレイ・送信済みアイテムの自動書き換えを有効にするか（既定: true）
        /// </summary>
        [DataMember]
        public bool EnableIncomingAndSentAutoProcess { get; set; } = true;

        /// <summary>
        /// 受信メールについて、差出人（From）だけでなく宛先（To/Cc）の表示名も書き換えるか（既定: false）
        /// </summary>
        [DataMember]
        public bool RewriteRecipientsOnIncomingMail { get; set; } = false;

        /// <summary>
        /// 処理対象外とするローカル（社内）ドメインの一覧（既定: kinki-sharyo.local）
        /// </summary>
        [DataMember]
        public List<string> ExcludedDomains { get; set; } = new List<string>
        {
            "kinki-sharyo.local"
        };

        /// <summary>
        /// 除外ドメインのサブドメイン（例: sub.kinki-sharyo.local）も除外対象に含めるか（既定: true）
        /// </summary>
        [DataMember]
        public bool IncludeSubdomains { get; set; } = true;

        /// <summary>
        /// 定期差分スキャンの間隔（分単位、既定: 5分）
        /// </summary>
        [DataMember]
        public int PeriodicScanIntervalMinutes { get; set; } = 5;

        /// <summary>
        /// 起動時・定期差分スキャンで遡る日数（既定: 7日間）
        /// </summary>
        [DataMember]
        public int AutoScanLookbackDays { get; set; } = 7;

        /// <summary>
        /// ログ出力の詳細度（既定: Info）
        /// </summary>
        [DataMember]
        public LogLevel LogVerbosity { get; set; } = LogLevel.Info;

        /// <summary>
        /// 設定のディープコピーを作成します。
        /// </summary>
        public AddInSettings Clone()
        {
            return new AddInSettings
            {
                EnableOutgoingProcess = this.EnableOutgoingProcess,
                IncludeBccInOutgoing = this.IncludeBccInOutgoing,
                ExpandDistributionListOnSend = this.ExpandDistributionListOnSend,
                EnableIncomingAndSentAutoProcess = this.EnableIncomingAndSentAutoProcess,
                RewriteRecipientsOnIncomingMail = this.RewriteRecipientsOnIncomingMail,
                ExcludedDomains = this.ExcludedDomains != null ? new List<string>(this.ExcludedDomains) : new List<string>(),
                IncludeSubdomains = this.IncludeSubdomains,
                PeriodicScanIntervalMinutes = this.PeriodicScanIntervalMinutes,
                AutoScanLookbackDays = this.AutoScanLookbackDays,
                LogVerbosity = this.LogVerbosity
            };
        }
    }
}
