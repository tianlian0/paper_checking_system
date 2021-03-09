using System;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using paper_checking.PaperCheck;

namespace paper_checking
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        /*
         * 获取当前UI的设置项的快照数据
         */
        private RunningEnv GetRunningEnv()
        {
            RunningEnv runningEnv = new RunningEnv(this);
            runningEnv.CheckData.CheckWay = cmbCheckWay.SelectedIndex;
            runningEnv.CheckData.CheckThreshold = int.Parse(txtCheckThreshold.Text);
            runningEnv.CheckData.Recover = chkRecover.Checked;
            runningEnv.CheckData.StatisTable = chkStatisTable.Checked;
            runningEnv.CheckData.ToCheckPaperPath = txtToCheckPaperPath.Text;
            runningEnv.CheckData.FinalReportPath = txtFinalReportPath.Text;
            runningEnv.CheckData.Blocklist = txtBlocklist.Text;

            runningEnv.LibraryData.PaperSourcePath = txtPaperSourcePath.Text;

            runningEnv.SettingData.CheckThreadCnt = int.Parse(txtCheckThreadCnt.Text);
            runningEnv.SettingData.ConvertThreadCnt = int.Parse(txtConvertThreadCnt.Text);
            runningEnv.CheckData.MinWords = int.Parse(txtMinWords.Text);
            runningEnv.CheckData.MinBytes = int.Parse(txtMinBytes.Text);
            runningEnv.SettingData.SuportPdf = chkSuportPdf.Checked;
            runningEnv.SettingData.SuportDoc = chkSuportDoc.Checked;
            runningEnv.SettingData.SuportDocx = chkSuportDocx.Checked;
            runningEnv.SettingData.SuportTxt = chkSuportTxt.Checked;

            return runningEnv;
        }

        /*
         * UI快照数据还原至UI
         */
        private void RestoreRunningEnv(RunningEnv runningEnv)
        {
            cmbCheckWay.SelectedIndex = runningEnv.CheckData.CheckWay;
            txtCheckThreshold.Text = runningEnv.CheckData.CheckThreshold.ToString();
            chkRecover.Checked = runningEnv.CheckData.Recover;
            chkStatisTable.Checked = runningEnv.CheckData.StatisTable;
            txtToCheckPaperPath.Text = runningEnv.CheckData.ToCheckPaperPath;
            txtFinalReportPath.Text = runningEnv.CheckData.FinalReportPath;
            txtBlocklist.Text = runningEnv.CheckData.Blocklist;

            txtPaperSourcePath.Text = runningEnv.LibraryData.PaperSourcePath;

            txtCheckThreadCnt.Text = runningEnv.SettingData.CheckThreadCnt.ToString();
            txtConvertThreadCnt.Text = runningEnv.SettingData.ConvertThreadCnt.ToString();
            txtMinWords.Text = runningEnv.CheckData.MinWords.ToString();
            txtMinBytes.Text = runningEnv.CheckData.MinBytes.ToString();
            chkSuportPdf.Checked = runningEnv.SettingData.SuportPdf;
            chkSuportDoc.Checked = runningEnv.SettingData.SuportDoc;
            chkSuportDocx.Checked = runningEnv.SettingData.SuportDocx;
            chkSuportTxt.Checked = runningEnv.SettingData.SuportTxt;
        }

        /*
         * 检查设置项，如果存在不符合条件的则恢复为默认值
         */
        private RunningEnv ConfigCheck(RunningEnv runningEnv)
        {
            if (runningEnv.CheckData.CheckThreshold >= 100 || runningEnv.CheckData.CheckThreshold < 1)
            {
                runningEnv.CheckData.CheckThreshold = 13;
            }
            if (runningEnv.CheckData.CheckWay != 0 && runningEnv.CheckData.CheckWay != 1)
            {
                runningEnv.CheckData.CheckWay = 0;
            }
            if (runningEnv.CheckData.MinBytes < 0)
            {
                runningEnv.CheckData.MinBytes = 0;
            }
            if(runningEnv.CheckData.MinWords < 0 || runningEnv.CheckData.MinWords > 99998)
            {
                runningEnv.CheckData.MinWords = 0;
            }
            if (runningEnv.SettingData.CheckThreadCnt >= 100 || runningEnv.SettingData.CheckThreadCnt < 1)
            {
                runningEnv.SettingData.CheckThreadCnt = 3;
            }
            if (runningEnv.SettingData.ConvertThreadCnt >= 100 || runningEnv.SettingData.ConvertThreadCnt < 1)
            {
                runningEnv.SettingData.ConvertThreadCnt = 2;
            }
            return runningEnv;
        }

        /*
         * 整体界面可用性控制
         */
        static public void SetComponentState(MainForm mainForm, bool state)
        {
            mainForm.BeginInvoke(new Action(() =>
                    {
                        mainForm.txtCheckThreshold.Enabled = state;
                        mainForm.groupBox1.Enabled = state;
                        mainForm.groupBox2.Enabled = state;
                        mainForm.groupBox4.Enabled = state;
                        mainForm.btnFinalReportPath.Enabled = state;
                        mainForm.btnStartChecking.Enabled = state;
                        mainForm.btnShowReportList.Enabled = state;
                        mainForm.btnExportReport.Enabled = state;
                        mainForm.btnToCheckPaperPath.Enabled = state;
                        mainForm.btnAddLibrary.Enabled = state;
                        mainForm.btnPaperSourcePath.Enabled = state;
                        mainForm.chkRecover.Enabled = state;
                        mainForm.chkStatisTable.Enabled = state;
                        mainForm.cmbCheckWay.Enabled = state;
                    }
                )
            );
        }

        /*
         * 论文库选项卡可用性控制
         */
        static public void SetPartComponentState(MainForm mainForm, bool state)
        {
            mainForm.BeginInvoke(new Action(() =>
                    {
                        mainForm.btnAddLibrary.Enabled = state;
                        mainForm.btnPaperSourcePath.Enabled = state;
                        mainForm.btnStartChecking.Enabled = state;
                        mainForm.groupBox2.Enabled = state;
                        mainForm.groupBox4.Enabled = state;
                    }
                )
            );
        }

        /*
         * 添加论文库
         */
        private void ButtonAddLibrary(object sender, EventArgs e)
        {
            //判断文件夹是否存在
            if (!Directory.Exists(txtPaperSourcePath.Text))
            {
                MessageBox.Show(this, "待添加的文件夹不存在！", "提示");
                return;
            }
            //启动添加任务
            PaperManager paperManager = new PaperManager(GetRunningEnv());
            Thread fileConvertThread = new Thread(new ThreadStart(paperManager.StartFileConvertStandalone));
            fileConvertThread.Start();
        }
        
        /*
         * 开始查重
         */ 
        private void ButtonStartChecking(object sender, EventArgs e)
        {
            //判断参数合法性
            if (int.Parse(txtCheckThreshold.Text) <= 0 || 
                int.Parse(txtCheckThreadCnt.Text) <= 0 || 
                int.Parse(txtConvertThreadCnt.Text) <= 0 ||
                int.Parse(txtMinBytes.Text) < 0 ||
                int.Parse(txtMinWords.Text) < 0)
            {
                MessageBox.Show(this, "线程数和查重阈值必须大于0，论文限制不可小于0！", "提示");
                return;
            }

            //判断所需文件夹是否已选择
            if (txtFinalReportPath.Text == "" || txtToCheckPaperPath.Text == "" || 
                !Directory.Exists(txtFinalReportPath.Text) || !Directory.Exists(txtToCheckPaperPath.Text))
            {
                MessageBox.Show(this, "待查论文文件夹或查重报告保存的文件夹不存在！", "提示");
                return;
            }

            //启动查重任务
            PaperManager paperManager = new PaperManager(GetRunningEnv());
            Thread checkThread = new Thread(new ThreadStart(paperManager.StartCheckPaper));
            checkThread.Start();
            btnForceStop.Focus();//重新设置焦点
        }

        /*
         * 查看查重报告列表按钮
         */
        private void ButtonShowReportList(object sender, EventArgs e)
        {
            ReportListForm reportListForm = null;
            try
            {
                //读取查重报告列表并展示
                reportListForm = new ReportListForm(RunningEnv.ProgramParam.ToCheckTxtPaperPath,
                                                        RunningEnv.ProgramParam.ReportPath,
                                                        RunningEnv.ProgramParam.ReportDataPath);
                reportListForm.ShowDialog();
            }
            catch {
                MessageBox.Show(this, "查重报告数据损坏，无法展示！", "错误");
            }
            finally
            {
                if (reportListForm != null)
                {
                    reportListForm.Dispose();
                }
            }
        }

        /*
         * 导出上一次查重报告按钮
         */
        private void ButtonExportReport(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PaperManager paperManager = new PaperManager(GetRunningEnv());
                paperManager.ExportReport();
            }
            dialog.Dispose();
        }

        /*
         * 阻止数字以外的字符输入
         */
        private void DigitFilterKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /*
         * 退出时写入配置文件
         */
        private void MainFormFormClosing(object sender, FormClosingEventArgs e)
        {
            if (tabControl1.TabPages.Count <= 1)
            {
                return;
            }

            //判断是否有正在进行的任务
            if (!Monitor.TryEnter(RunningEnv.EnvRunningLock))
            {
                e.Cancel = true;//取消关闭
                MessageBox.Show("有任务正在进行，请等待任务结束！", "提示");
                return;
            }

            //锁释放
            Monitor.Exit(RunningEnv.EnvRunningLock);
            //保存配置项到配置文件
            Utils.SaveConfig(GetRunningEnv());
        }

        /*
         * 启动时读取配置文件
         */
        private void MainFormLoad(object sender, EventArgs e)
        {
            lblVersion.Text = Application.ProductVersion;
            cmbCheckWay.SelectedIndex = 0;
            //如果存在配置文件
            if (File.Exists("config.ini"))
            {
                //读取设置项合法性检查并恢复至UI
                RestoreRunningEnv(ConfigCheck(Utils.ReadConfig(this)));
            }
            else
            {
                //恢复默认设置
                ButtonRestoreDefault1(sender, e);
                ButtonRestoreDefault2(sender, e);
            }
            //创建初始文件夹
            Utils.FolderCheck();
        }

        /*
         * 选择查重报告保存路径
         */
        private void ButtonSelectFinalReportPath(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtFinalReportPath.Text = dialog.SelectedPath;
            }
            dialog.Dispose();
        }

        /*
         * 重置
         */
        private void ButtonReset(object sender, EventArgs e)
        {
            Utils.ResetProgram();
        }

        /*
         * 强制退出
         */
        private void ButtonForceStop(object sender, EventArgs e)
        {
            if (!Monitor.TryEnter(RunningEnv.EnvRunningLock))
            {
                DialogResult dr = MessageBox.Show("有任务正在进行，强制退出有可能丢失数据，是否继续本次强制退出？", "提示", MessageBoxButtons.OKCancel);
                if (dr != DialogResult.OK)
                {
                    return;
                }
                else
                {
                    Utils.ForceStop();
                }
            }
            Monitor.Exit(RunningEnv.EnvRunningLock);
            Application.Exit();
        }

        /*
         * 选择待查论文路径
         */
        private void ButtonSelectToCheckPaperPath(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtToCheckPaperPath.Text = dialog.SelectedPath;
            }
            dialog.Dispose();
        }

        /*
         * 选择待添加论文库的论文文件夹
         */
        private void ButtonSelectPaperSourcePath(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtPaperSourcePath.Text = dialog.SelectedPath;
            }
            dialog.Dispose();
        }


        private void ButtonShowLicence(object sender, EventArgs e)
        {
            new Licence().ShowDialog();
        }

        /*
         * 恢复默认值
         */
        private void ButtonRestoreDefault1(object sender, EventArgs e)
        {
            //获取CPU逻辑核心数并-2
            int processorCount = Environment.ProcessorCount - 2;
            if (processorCount > 1)
            {
                txtCheckThreadCnt.Text = processorCount.ToString();
                txtConvertThreadCnt.Text = processorCount.ToString();
            }
            else
            {
                txtCheckThreadCnt.Text = "1";
                txtConvertThreadCnt.Text = "1";
            }
        }

        /*
         * 恢复默认值
         */
        private void ButtonRestoreDefault2(object sender, EventArgs e)
        {
            txtMinWords.Text = "0";
            txtMinBytes.Text = "0";
        }

        private void TxtBlocklist_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 || e.KeyChar == 10 || e.KeyChar == 64)
                e.Handled = true;
        }

        private void BtnManageLibrary_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Application.StartupPath + Path.DirectorySeparatorChar + RunningEnv.ProgramParam.TxtPaperSourcePath);
        }
    }
   
}
