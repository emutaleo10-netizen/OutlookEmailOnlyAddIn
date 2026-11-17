using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using OutlookEmailOnlyAddIn.Logging;

namespace OutlookEmailOnlyAddIn.UI
{
    /// <summary>
    /// Outlook のリボンに「表示名除去アドイン」グループ（設定ボタン・一括処理ボタン）を追加するクラス
    /// </summary>
    [ComVisible(true)]
    public class AddInRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            // メインウィンドウ（Microsoft.Outlook.Explorer）の「ホーム」タブに表示する
            if (string.Equals(ribbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
            {
                return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""EmailOnlyMainGroup"" label=""表示名除去"">
          <button id=""BtnOpenSettings""
                  label=""表示名除去設定""
                  size=""large""
                  imageMso=""FileProperties""
                  onAction=""OnOpenSettingsClick""
                  screentip=""メールアドレス統一表示の設定""
                  supertip=""送信・受信メールの表示名除去機能のON/OFF、除外ドメイン（@kinki-sharyo.local等）、Bcc対象設定、ログ出力設定を変更します。"" />
          <button id=""BtnOpenBatchProcess""
                  label=""既存メール一括処理""
                  size=""large""
                  imageMso=""SynchronizeAll""
                  onAction=""OnOpenBatchProcessClick""
                  screentip=""フォルダ単位の一括書き換え""
                  supertip=""指定したフォルダ内の既存メールに対して、差出人名・宛先名のメールアドレス化を一括適用します。"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
            }

            return null;
        }

        /// <summary>
        /// リボンロード時のコールバック
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            _ribbon?.Invalidate();
        }

        /// <summary>
        /// 「表示名除去設定」ボタン押下時のコールバック
        /// </summary>
        public void OnOpenSettingsClick(Office.IRibbonControl control)
        {
            try
            {
                using (var form = new SettingsForm())
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("設定画面の表示中にエラーが発生しました。", ex);
                MessageBox.Show(
                    "設定画面を開けませんでした。\n" + ex.Message,
                    "表示名除去アドイン",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「既存メール一括処理」ボタン押下時のコールバック
        /// </summary>
        public void OnOpenBatchProcessClick(Office.IRibbonControl control)
        {
            try
            {
                OpenBatchProcessDialog();
            }
            catch (Exception ex)
            {
                FileLogger.Error("一括処理画面の表示中にエラーが発生しました。", ex);
                MessageBox.Show(
                    "一括処理画面を開けませんでした。\n" + ex.Message,
                    "表示名除去アドイン",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 一括処理ダイアログを起動します（設定画面からも呼び出し可能）。
        /// </summary>
        internal static void OpenBatchProcessDialog(IWin32Window owner = null)
        {
            using (var form = new BatchProcessForm(Globals.ThisAddIn.Application))
            {
                if (owner != null)
                {
                    form.ShowDialog(owner);
                }
                else
                {
                    form.ShowDialog();
                }
            }
        }
    }
}
