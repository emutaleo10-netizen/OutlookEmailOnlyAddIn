using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using OutlookEmailOnlyAddIn.Core;
using OutlookEmailOnlyAddIn.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailOnlyAddIn.UI
{
    /// <summary>
    /// フォルダ単位の既存メール一括処理（対象フォルダ選択・進捗表示・中断/再開）ダイアログフォーム
    /// </summary>
    public class BatchProcessForm : Form
    {
        private readonly Outlook.Application _application;

        private TreeView _tvFolders;
        private CheckBox _chkIncludeSubfolders;
        private ComboBox _cmbProcessMode;
        private CheckBox _chkForceReprocess;

        private ProgressBar _progressBar;
        private Label _lblStatus;
        private Label _lblStats;

        private Button _btnStart;
        private Button _btnCancelProcess;
        private Button _btnClose;

        private CancellationTokenSource _cts;
        private bool _isRunning;
        private bool _suppressTreeCheckEvent;

        private sealed class FolderNodeTag
        {
            public string FolderPath { get; set; }
            public string EntryId { get; set; }
            public string StoreId { get; set; }
            public bool IsSentDefaultFolder { get; set; }
        }

        public BatchProcessForm(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            InitializeComponents();
            PopulateFolderTree();
        }

        private void InitializeComponents()
        {
            this.Text = "表示名除去アドイン - 既存メールの一括処理";
            this.Size = new Size(620, 580);
            this.MinimumSize = new Size(620, 580);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            int y = 14;

            var lblInstruction = new Label
            {
                Text = "一括処理を実行する対象フォルダにチェックを入れてください（中断しても次回未処理分から再開できます）:",
                Location = new Point(14, y),
                Size = new Size(580, 20)
            };
            this.Controls.Add(lblInstruction);
            y += 24;

            _tvFolders = new TreeView
            {
                Location = new Point(14, y),
                Size = new Size(576, 250),
                CheckBoxes = true,
                HideSelection = false
            };
            _tvFolders.AfterCheck += TvFolders_AfterCheck;
            this.Controls.Add(_tvFolders);
            y += _tvFolders.Height + 10;

            // --- オプション設定グループ ---
            var grpOptions = new GroupBox
            {
                Text = "一括処理オプション",
                Location = new Point(14, y),
                Size = new Size(576, 112)
            };

            _chkIncludeSubfolders = new CheckBox
            {
                Text = "親フォルダにチェックを入れた際、配下のサブフォルダも自動的に選択する",
                Location = new Point(16, 22),
                Size = new Size(540, 22),
                Checked = true
            };

            var lblMode = new Label
            {
                Text = "書き換えモード:",
                Location = new Point(16, 52),
                Size = new Size(95, 20)
            };

            _cmbProcessMode = new ComboBox
            {
                Location = new Point(115, 49),
                Size = new Size(440, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbProcessMode.Items.Add("自動判別（送信済みアイテムは宛先名、それ以外のフォルダは差出人名を書き換え）");
            _cmbProcessMode.Items.Add("差出人（From）の表示名のみを書き換える（受信メール向け）");
            _cmbProcessMode.Items.Add("宛先（To / Cc）の表示名を書き換える（送信済みメール向け）");
            _cmbProcessMode.SelectedIndex = 0;

            _chkForceReprocess = new CheckBox
            {
                Text = "既に処理済みマークが付いているメールも再チェックする（除外ドメイン設定を変更した後に使用）",
                Location = new Point(16, 80),
                Size = new Size(545, 22),
                Checked = false
            };

            grpOptions.Controls.Add(_chkIncludeSubfolders);
            grpOptions.Controls.Add(lblMode);
            grpOptions.Controls.Add(_cmbProcessMode);
            grpOptions.Controls.Add(_chkForceReprocess);
            this.Controls.Add(grpOptions);
            y += grpOptions.Height + 12;

            // --- 進捗表示エリア ---
            _lblStatus = new Label
            {
                Text = "待機中（対象フォルダを選択して「一括処理開始」を押してください）",
                Location = new Point(14, y),
                Size = new Size(576, 20),
                AutoEllipsis = true
            };
            this.Controls.Add(_lblStatus);
            y += 22;

            _progressBar = new ProgressBar
            {
                Location = new Point(14, y),
                Size = new Size(576, 24),
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            this.Controls.Add(_progressBar);
            y += 30;

            _lblStats = new Label
            {
                Text = "進捗: 0 / 0 件  (書き換え: 0, 処理済みスキップ: 0, 変更なし: 0, 対象外/エラー: 0)",
                Location = new Point(14, y),
                Size = new Size(576, 20)
            };
            this.Controls.Add(_lblStats);
            y += 28;

            // --- 下部ボタン ---
            _btnStart = new Button
            {
                Text = "一括処理開始(&S)",
                Location = new Point(236, y),
                Size = new Size(125, 32)
            };
            _btnStart.Click += BtnStart_Click;

            _btnCancelProcess = new Button
            {
                Text = "中断(&C)",
                Location = new Point(370, y),
                Size = new Size(105, 32),
                Enabled = false
            };
            _btnCancelProcess.Click += BtnCancelProcess_Click;

            _btnClose = new Button
            {
                Text = "閉じる",
                Location = new Point(485, y),
                Size = new Size(105, 32)
            };
            _btnClose.Click += (s, e) => this.Close();

            this.Controls.Add(_btnStart);
            this.Controls.Add(_btnCancelProcess);
            this.Controls.Add(_btnClose);
        }

        /// <summary>
        /// Outlook プロファイル内の全ストア・メールフォルダを TreeView に読み込みます。
        /// 既定で「受信トレイ」と「送信済みアイテム」にチェックを入れた状態にします。
        /// </summary>
        private void PopulateFolderTree()
        {
            _tvFolders.BeginUpdate();
            _tvFolders.Nodes.Clear();

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
                    Outlook.MAPIFolder rootFolder = null;
                    try
                    {
                        store = stores[i];
                        string inboxId = SafeGetDefaultFolderEntryId(store, Outlook.OlDefaultFolders.olFolderInbox);
                        string sentId = SafeGetDefaultFolderEntryId(store, Outlook.OlDefaultFolders.olFolderSentMail);

                        rootFolder = store.GetRootFolder();
                        if (rootFolder != null)
                        {
                            TreeNode storeNode = CreateFolderTreeNode(rootFolder, inboxId, sentId, isRoot: true);
                            _tvFolders.Nodes.Add(storeNode);
                            storeNode.Expand();
                        }
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Debug($"ストアのフォルダツリー構築をスキップしました: {ex.Message}");
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(rootFolder);
                        MailAddressHelper.ReleaseCom(store);
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("フォルダツリーの構築中にエラーが発生しました。", ex);
            }
            finally
            {
                MailAddressHelper.ReleaseCom(stores);
                MailAddressHelper.ReleaseCom(session);
                _tvFolders.EndUpdate();
            }
        }

        private TreeNode CreateFolderTreeNode(
            Outlook.MAPIFolder folder,
            string inboxEntryId,
            string sentEntryId,
            bool isRoot)
        {
            string name = folder.Name ?? "(名称なし)";
            string entryId = folder.EntryID;
            string storeId = folder.StoreID;
            string folderPath = folder.FolderPath ?? name;

            bool isSent = !string.IsNullOrEmpty(sentEntryId) &&
                          string.Equals(entryId, sentEntryId, StringComparison.OrdinalIgnoreCase);
            bool isInbox = !string.IsNullOrEmpty(inboxEntryId) &&
                           string.Equals(entryId, inboxEntryId, StringComparison.OrdinalIgnoreCase);

            var node = new TreeNode(name)
            {
                Tag = new FolderNodeTag
                {
                    FolderPath = folderPath,
                    EntryId = entryId,
                    StoreId = storeId,
                    IsSentDefaultFolder = isSent
                },
                Checked = isInbox || isSent
            };

            Outlook.Folders subFolders = null;
            try
            {
                subFolders = folder.Folders;
                int count = subFolders.Count;

                for (int i = 1; i <= count; i++)
                {
                    Outlook.MAPIFolder sub = null;
                    try
                    {
                        sub = subFolders[i];
                        // メールアイテム用フォルダ（olMailItem）またはルート配下のみ表示（予定表・連絡先・タスク等は除外）
                        if (sub.DefaultItemType == Outlook.OlItemType.olMailItem)
                        {
                            // 削除済みアイテム・送信トレイ・下書き・迷惑メールは初期ツリーから誤選択を防ぐためチェックなし
                            TreeNode childNode = CreateFolderTreeNode(sub, inboxEntryId, sentEntryId, isRoot: false);
                            node.Nodes.Add(childNode);
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        MailAddressHelper.ReleaseCom(sub);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                MailAddressHelper.ReleaseCom(subFolders);
            }

            if (isInbox || isSent)
            {
                node.Expand();
            }

            return node;
        }

        private static string SafeGetDefaultFolderEntryId(Outlook.Store store, Outlook.OlDefaultFolders folderType)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = store.GetDefaultFolder(folderType);
                return folder?.EntryID;
            }
            catch
            {
                return null;
            }
            finally
            {
                MailAddressHelper.ReleaseCom(folder);
            }
        }

        private void TvFolders_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_suppressTreeCheckEvent || !_chkIncludeSubfolders.Checked)
            {
                return;
            }

            if (e.Action == TreeViewAction.Unknown)
            {
                return;
            }

            try
            {
                _suppressTreeCheckEvent = true;
                SetChildrenChecked(e.Node, e.Node.Checked);
            }
            finally
            {
                _suppressTreeCheckEvent = false;
            }
        }

        private static void SetChildrenChecked(TreeNode parentNode, bool isChecked)
        {
            foreach (TreeNode child in parentNode.Nodes)
            {
                child.Checked = isChecked;
                SetChildrenChecked(child, isChecked);
            }
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning)
            {
                return;
            }

            int modeIndex = _cmbProcessMode.SelectedIndex;
            var targets = new List<BatchFolderTarget>();
            foreach (TreeNode rootNode in _tvFolders.Nodes)
            {
                CollectCheckedTargets(rootNode, modeIndex, parentIsSent: false, targets);
            }

            if (targets.Count == 0)
            {
                MessageBox.Show(
                    "一括処理を行うフォルダにチェックを入れてください。",
                    "フォルダ未選択",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _isRunning = true;
            SetUiRunningState(true);
            _cts = new CancellationTokenSource();

            var progress = new Progress<BatchProgressInfo>(UpdateProgressUi);
            var processor = new BatchProcessor(_application);

            try
            {
                BatchProgressInfo finalStats = await processor.ExecuteAsync(
                    targets,
                    forceReprocess: _chkForceReprocess.Checked,
                    progress: progress,
                    cancellationToken: _cts.Token);

                UpdateProgressUi(finalStats);
                _lblStatus.Text = "一括処理が完了しました。";

                MessageBox.Show(
                    $"一括処理が完了しました。\n\n" +
                    $"対象メール総数: {finalStats.TotalCount} 件\n" +
                    $"表示名を書き換えた件数: {finalStats.ModifiedCount} 件\n" +
                    $"変更なし（除外ドメイン等）: {finalStats.UnchangedCount} 件\n" +
                    $"処理済みのためスキップ: {finalStats.AlreadyProcessedCount} 件\n" +
                    $"対象外またはエラー: {finalStats.SkippedOrErrorCount} 件",
                    "一括処理完了",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                _lblStatus.Text = "一括処理を中断しました（次回実行時は未処理のメールから再開できます）。";
                MessageBox.Show(
                    "一括処理を中断しました。\n既に完了したメールには処理済みマークが記録されているため、次回実行時は残りのメールから再開されます。",
                    "一括処理中断",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "エラーが発生しました: " + ex.Message;
                FileLogger.Error("一括処理の実行中に予期しない例外が発生しました。", ex);
                MessageBox.Show(
                    "一括処理中にエラーが発生しました。\n" + ex.Message,
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _cts?.Dispose();
                _cts = null;
                _isRunning = false;
                SetUiRunningState(false);
            }
        }

        private void CollectCheckedTargets(
            TreeNode node,
            int modeIndex,
            bool parentIsSent,
            List<BatchFolderTarget> output)
        {
            var tag = node.Tag as FolderNodeTag;
            bool isSent = parentIsSent || (tag != null && tag.IsSentDefaultFolder);

            if (node.Checked && tag != null && !string.IsNullOrEmpty(tag.EntryId))
            {
                bool treatAsSent;
                if (modeIndex == 1)
                {
                    treatAsSent = false; // 強制的に差出人書き換え
                }
                else if (modeIndex == 2)
                {
                    treatAsSent = true;  // 強制的に宛先書き換え
                }
                else
                {
                    treatAsSent = isSent; // 自動判別
                }

                output.Add(new BatchFolderTarget
                {
                    FolderPath = tag.FolderPath,
                    EntryId = tag.EntryId,
                    StoreId = tag.StoreId,
                    TreatAsSentFolder = treatAsSent
                });
            }

            foreach (TreeNode child in node.Nodes)
            {
                CollectCheckedTargets(child, modeIndex, isSent, output);
            }
        }

        private void UpdateProgressUi(BatchProgressInfo info)
        {
            if (this.IsDisposed)
            {
                return;
            }

            _lblStatus.Text = $"処理中フォルダ: {info.CurrentFolderPath}";

            if (info.TotalCount > 0)
            {
                int percent = (int)Math.Min(100, Math.Round((double)info.ProcessedCount * 100.0 / info.TotalCount));
                _progressBar.Value = percent;
            }
            else
            {
                _progressBar.Value = 0;
            }

            _lblStats.Text =
                $"進捗: {info.ProcessedCount} / {info.TotalCount} 件  " +
                $"(書き換え: {info.ModifiedCount}, 処理済みスキップ: {info.AlreadyProcessedCount}, " +
                $"変更なし: {info.UnchangedCount}, 対象外/エラー: {info.SkippedOrErrorCount})";
        }

        private void BtnCancelProcess_Click(object sender, EventArgs e)
        {
            if (_isRunning && _cts != null && !_cts.IsCancellationRequested)
            {
                _btnCancelProcess.Enabled = false;
                _lblStatus.Text = "中断要求を受け付けました。現在のチャンク完了後に停止します...";
                _cts.Cancel();
            }
        }

        private void SetUiRunningState(bool running)
        {
            _tvFolders.Enabled = !running;
            _chkIncludeSubfolders.Enabled = !running;
            _cmbProcessMode.Enabled = !running;
            _chkForceReprocess.Enabled = !running;
            _btnStart.Enabled = !running;
            _btnCancelProcess.Enabled = running;
            _btnClose.Enabled = !running;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_isRunning)
            {
                e.Cancel = true;
                MessageBox.Show(
                    "一括処理の実行中はウィンドウを閉じられません。先に「中断」ボタンを押してください。",
                    "実行中",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            base.OnFormClosing(e);
        }
    }
}
