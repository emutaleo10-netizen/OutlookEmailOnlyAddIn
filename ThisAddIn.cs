using System;
using Office = Microsoft.Office.Core;
using OutlookEmailOnlyAddIn.Configuration;
using OutlookEmailOnlyAddIn.Core;
using OutlookEmailOnlyAddIn.Logging;
using OutlookEmailOnlyAddIn.UI;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn
{
    /// <summary>
    /// VSTO アドインのエントリポイントクラス
    /// </summary>
    public partial class ThisAddIn
    {
        private FolderMonitor _folderMonitor;

        /// <summary>
        /// カスタムリボン（表示名除去設定・既存メール一括処理ボタン）を Outlook に登録します。
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AddInRibbon();
        }

        /// <summary>
        /// アドイン起動時の初期化処理
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                AddInSettings settings = SettingsManager.Current;
                FileLogger.Info("========================================");
                FileLogger.Info("OutlookEmailOnlyAddIn を起動しました。");
                FileLogger.Info($"設定ファイル: {SettingsManager.SettingsFilePath}");
                FileLogger.Info($"送信処理: {(settings.EnableOutgoingProcess ? "ON" : "OFF")}, Bcc対象: {(settings.IncludeBccInOutgoing ? "ON" : "OFF")}, 受信・送信済み自動処理: {(settings.EnableIncomingAndSentAutoProcess ? "ON" : "OFF")}");
                FileLogger.Info($"除外ドメイン: {string.Join(", ", settings.ExcludedDomains)} (サブドメイン含む: {settings.IncludeSubdomains})");

                // ステップ1: 送信イベント（ItemSend）の登録
                this.Application.ItemSend += Application_ItemSend;

                // ステップ2: 受信トレイ・送信済みアイテムの監視および起動時・定期差分スキャンの開始
                _folderMonitor = new FolderMonitor(this.Application);
                _folderMonitor.Start();
            }
            catch (Exception ex)
            {
                FileLogger.Error("アドイン起動処理中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// アドイン終了時のクリーンアップ処理
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                this.Application.ItemSend -= Application_ItemSend;

                if (_folderMonitor != null)
                {
                    _folderMonitor.Dispose();
                    _folderMonitor = null;
                }

                FileLogger.Info("OutlookEmailOnlyAddIn を終了しました。");
            }
            catch (Exception ex)
            {
                FileLogger.Error("アドイン終了処理中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// メール送信時イベントハンドラ
        /// </summary>
        /// <param name="Item">送信されるアイテム</param>
        /// <param name="Cancel">送信をキャンセルする場合は true（当アドインでは常に false を維持）</param>
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            // 例外が発生しても Cancel は書き換えずメール送信を妨げない
            OutgoingMailProcessor.ProcessOutgoingItem(Item);
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
