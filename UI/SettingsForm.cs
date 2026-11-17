using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OutlookEmailOnlyAddIn.Configuration;
using OutlookEmailOnlyAddIn.Logging;

namespace OutlookEmailOnlyAddIn.UI
{
    /// <summary>
    /// 表示名除去アドインの各種設定を行うダイアログフォーム
    /// </summary>
    public class SettingsForm : Form
    {
        private CheckBox _chkEnableOutgoing;
        private CheckBox _chkIncludeBcc;
        private CheckBox _chkExpandDistList;

        private CheckBox _chkEnableIncomingAuto;
        private CheckBox _chkRewriteIncomingRecipients;
        private NumericUpDown _numScanInterval;
        private NumericUpDown _numLookbackDays;

        private ListBox _lstExcludedDomains;
        private TextBox _txtNewDomain;
        private Button _btnAddDomain;
        private Button _btnRemoveDomain;
        private CheckBox _chkIncludeSubdomains;

        private ComboBox _cmbLogLevel;
        private Button _btnOpenLogFolder;
        private Button _btnOpenBatchForm;

        private Button _btnSave;
        private Button _btnCancel;

        public SettingsForm()
        {
            InitializeComponents();
            LoadCurrentSettings();
        }

        private void InitializeComponents()
        {
            this.Text = "表示名除去アドイン - 設定";
            this.Size = new Size(560, 630);
            this.MinimumSize = new Size(560, 630);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            int y = 14;

            // --- 1. 送信メール設定グループ ---
            var grpOutgoing = new GroupBox
            {
                Text = "1. 送信メール設定",
                Location = new Point(14, y),
                Size = new Size(516, 108)
            };

            _chkEnableOutgoing = new CheckBox
            {
                Text = "送信時に宛先（To / Cc）の表示名を除去し、メールアドレスのみにする",
                Location = new Point(16, 24),
                Size = new Size(480, 22),
                AutoSize = false
            };

            _chkIncludeBcc = new CheckBox
            {
                Text = "Bcc の宛先も表示名除去の対象にする",
                Location = new Point(34, 50),
                Size = new Size(460, 22),
                AutoSize = false
            };

            _chkExpandDistList = new CheckBox
            {
                Text = "連絡先グループ（配布リスト）を個別メンバーに展開して処理する",
                Location = new Point(34, 76),
                Size = new Size(460, 22),
                AutoSize = false
            };

            grpOutgoing.Controls.Add(_chkEnableOutgoing);
            grpOutgoing.Controls.Add(_chkIncludeBcc);
            grpOutgoing.Controls.Add(_chkExpandDistList);
            this.Controls.Add(grpOutgoing);

            y += grpOutgoing.Height + 10;

            // --- 2. 受信トレイ・送信済みアイテム設定グループ ---
            var grpIncoming = new GroupBox
            {
                Text = "2. 受信トレイ / 送信済みアイテムの自動書き換え設定",
                Location = new Point(14, y),
                Size = new Size(516, 118)
            };

            _chkEnableIncomingAuto = new CheckBox
            {
                Text = "新着受信メールの差出人名・送信済みメールの宛先名を自動でメールアドレス表示にする",
                Location = new Point(16, 24),
                Size = new Size(490, 22),
                AutoSize = false
            };

            _chkRewriteIncomingRecipients = new CheckBox
            {
                Text = "受信メール内の宛先（To / Cc）の表示名も併せてメールアドレスに書き換える",
                Location = new Point(34, 50),
                Size = new Size(470, 22),
                AutoSize = false
            };

            var lblInterval = new Label
            {
                Text = "定期差分スキャン間隔（分）:",
                Location = new Point(34, 82),
                Size = new Size(165, 20)
            };

            _numScanInterval = new NumericUpDown
            {
                Location = new Point(200, 79),
                Size = new Size(60, 23),
                Minimum = 1,
                Maximum = 120,
                Value = 5
            };

            var lblLookback = new Label
            {
                Text = "自動スキャン遡り日数（日）:",
                Location = new Point(280, 82),
                Size = new Size(160, 20)
            };

            _numLookbackDays = new NumericUpDown
            {
                Location = new Point(442, 79),
                Size = new Size(55, 23),
                Minimum = 1,
                Maximum = 90,
                Value = 7
            };

            grpIncoming.Controls.Add(_chkEnableIncomingAuto);
            grpIncoming.Controls.Add(_chkRewriteIncomingRecipients);
            grpIncoming.Controls.Add(lblInterval);
            grpIncoming.Controls.Add(_numScanInterval);
            grpIncoming.Controls.Add(lblLookback);
            grpIncoming.Controls.Add(_numLookbackDays);
            this.Controls.Add(grpIncoming);

            y += grpIncoming.Height + 10;

            // --- 3. 除外ローカルドメイン設定グループ ---
            var grpDomains = new GroupBox
            {
                Text = "3. 除外ルール（社内・ローカルドメイン指定）",
                Location = new Point(14, y),
                Size = new Size(516, 195)
            };

            var lblDomainDesc = new Label
            {
                Text = "以下のドメインに該当するアドレスは表示名を除去せず、従来の表示名を維持します:",
                Location = new Point(16, 22),
                Size = new Size(480, 18)
            };

            _lstExcludedDomains = new ListBox
            {
                Location = new Point(16, 44),
                Size = new Size(365, 92)
            };

            _btnRemoveDomain = new Button
            {
                Text = "選択削除(&D)",
                Location = new Point(392, 44),
                Size = new Size(108, 28)
            };
            _btnRemoveDomain.Click += BtnRemoveDomain_Click;

            _txtNewDomain = new TextBox
            {
                Location = new Point(16, 140),
                Size = new Size(365, 23)
            };

            _btnAddDomain = new Button
            {
                Text = "ドメイン追加(&A)",
                Location = new Point(392, 138),
                Size = new Size(108, 26)
            };
            _btnAddDomain.Click += BtnAddDomain_Click;

            _chkIncludeSubdomains = new CheckBox
            {
                Text = "サブドメインも除外対象に含める（例: sub.kinki-sharyo.local も除外）",
                Location = new Point(16, 168),
                Size = new Size(480, 22),
                AutoSize = false
            };

            grpDomains.Controls.Add(lblDomainDesc);
            grpDomains.Controls.Add(_lstExcludedDomains);
            grpDomains.Controls.Add(_btnRemoveDomain);
            grpDomains.Controls.Add(_txtNewDomain);
            grpDomains.Controls.Add(_btnAddDomain);
            grpDomains.Controls.Add(_chkIncludeSubdomains);
            this.Controls.Add(grpDomains);

            y += grpDomains.Height + 10;

            // --- 4. ログ設定・既存メール一括処理グループ ---
            var grpMisc = new GroupBox
            {
                Text = "4. ログ出力・既存メールの一括処理",
                Location = new Point(14, y),
                Size = new Size(516, 86)
            };

            var lblLogLevel = new Label
            {
                Text = "ログ詳細度:",
                Location = new Point(16, 26),
                Size = new Size(76, 20)
            };

            _cmbLogLevel = new ComboBox
            {
                Location = new Point(96, 22),
                Size = new Size(165, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbLogLevel.Items.Add("Error（エラーのみ）");
            _cmbLogLevel.Items.Add("Info（標準・処理件数）");
            _cmbLogLevel.Items.Add("Debug（詳細トレース）");

            _btnOpenLogFolder = new Button
            {
                Text = "ログフォルダを開く(&L)",
                Location = new Point(276, 20),
                Size = new Size(150, 28)
            };
            _btnOpenLogFolder.Click += BtnOpenLogFolder_Click;

            _btnOpenBatchForm = new Button
            {
                Text = "既存メールの一括処理画面を開く(&B)...",
                Location = new Point(16, 52),
                Size = new Size(270, 26)
            };
            _btnOpenBatchForm.Click += BtnOpenBatchForm_Click;

            grpMisc.Controls.Add(lblLogLevel);
            grpMisc.Controls.Add(_cmbLogLevel);
            grpMisc.Controls.Add(_btnOpenLogFolder);
            grpMisc.Controls.Add(_btnOpenBatchForm);
            this.Controls.Add(grpMisc);

            y += grpMisc.Height + 12;

            // --- 下部ボタン ---
            _btnSave = new Button
            {
                Text = "保存(&S)",
                Location = new Point(324, y),
                Size = new Size(98, 30),
                DialogResult = DialogResult.OK
            };
            _btnSave.Click += BtnSave_Click;

            _btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(432, y),
                Size = new Size(98, 30),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.Add(_btnSave);
            this.Controls.Add(_btnCancel);
            this.AcceptButton = _btnSave;
            this.CancelButton = _btnCancel;
        }

        private void LoadCurrentSettings()
        {
            AddInSettings s = SettingsManager.Current;

            _chkEnableOutgoing.Checked = s.EnableOutgoingProcess;
            _chkIncludeBcc.Checked = s.IncludeBccInOutgoing;
            _chkExpandDistList.Checked = s.ExpandDistributionListOnSend;

            _chkEnableIncomingAuto.Checked = s.EnableIncomingAndSentAutoProcess;
            _chkRewriteIncomingRecipients.Checked = s.RewriteRecipientsOnIncomingMail;
            _numScanInterval.Value = Math.Max(_numScanInterval.Minimum, Math.Min(_numScanInterval.Maximum, s.PeriodicScanIntervalMinutes));
            _numLookbackDays.Value = Math.Max(_numLookbackDays.Minimum, Math.Min(_numLookbackDays.Maximum, s.AutoScanLookbackDays));

            _lstExcludedDomains.Items.Clear();
            if (s.ExcludedDomains != null)
            {
                foreach (string d in s.ExcludedDomains)
                {
                    _lstExcludedDomains.Items.Add(d);
                }
            }

            _chkIncludeSubdomains.Checked = s.IncludeSubdomains;

            int logIdx = (int)s.LogVerbosity;
            if (logIdx < 0 || logIdx >= _cmbLogLevel.Items.Count)
            {
                logIdx = (int)LogLevel.Info;
            }
            _cmbLogLevel.SelectedIndex = logIdx;
        }

        private void BtnAddDomain_Click(object sender, EventArgs e)
        {
            string input = (_txtNewDomain.Text ?? string.Empty).Trim().TrimStart('@').TrimStart('.');
            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("追加するドメイン名（例: kinki-sharyo.local）を入力してください。", "ドメイン追加", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (input.Contains(" ") || input.Contains("@") || !input.Contains("."))
            {
                MessageBox.Show("有効なドメイン形式（例: kinki-sharyo.local）で入力してください。", "ドメイン追加", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (object item in _lstExcludedDomains.Items)
            {
                if (string.Equals(item.ToString(), input, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("指定されたドメインは既に一覧に登録されています。", "ドメイン追加", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            _lstExcludedDomains.Items.Add(input);
            _txtNewDomain.Clear();
            _txtNewDomain.Focus();
        }

        private void BtnRemoveDomain_Click(object sender, EventArgs e)
        {
            int idx = _lstExcludedDomains.SelectedIndex;
            if (idx >= 0)
            {
                _lstExcludedDomains.Items.RemoveAt(idx);
            }
        }

        private void BtnOpenLogFolder_Click(object sender, EventArgs e)
        {
            try
            {
                string logDir = FileLogger.LogDirectory;
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                Process.Start("explorer.exe", logDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ログフォルダを開けませんでした: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOpenBatchForm_Click(object sender, EventArgs e)
        {
            AddInRibbon.OpenBatchProcessDialog(this);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var domains = new List<string>();
            foreach (object item in _lstExcludedDomains.Items)
            {
                if (item != null)
                {
                    domains.Add(item.ToString());
                }
            }

            var updated = new AddInSettings
            {
                EnableOutgoingProcess = _chkEnableOutgoing.Checked,
                IncludeBccInOutgoing = _chkIncludeBcc.Checked,
                ExpandDistributionListOnSend = _chkExpandDistList.Checked,
                EnableIncomingAndSentAutoProcess = _chkEnableIncomingAuto.Checked,
                RewriteRecipientsOnIncomingMail = _chkRewriteIncomingRecipients.Checked,
                PeriodicScanIntervalMinutes = (int)_numScanInterval.Value,
                AutoScanLookbackDays = (int)_numLookbackDays.Value,
                ExcludedDomains = domains,
                IncludeSubdomains = _chkIncludeSubdomains.Checked,
                LogVerbosity = (LogLevel)_cmbLogLevel.SelectedIndex
            };

            SettingsManager.Save(updated);
            FileLogger.Info("設定画面からユーザー設定が更新・保存されました。");
        }
    }
}
