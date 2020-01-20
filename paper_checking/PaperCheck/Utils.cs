using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace paper_checking.PaperCheck
{
    static class Utils
    {
        /*
         * 创建存放论文和论文库的文件夹
         */
        static public bool FolderCheck()
        {
            try
            {
                if (!Directory.Exists(RunningEnv.ProgramParam.TxtPaperSourcePath))
                    Directory.CreateDirectory(RunningEnv.ProgramParam.TxtPaperSourcePath);
                if (!Directory.Exists(RunningEnv.ProgramParam.ToCheckTxtPaperPath))
                    Directory.CreateDirectory(RunningEnv.ProgramParam.ToCheckTxtPaperPath);
                if (!Directory.Exists(RunningEnv.ProgramParam.ReportPath))
                    Directory.CreateDirectory(RunningEnv.ProgramParam.ReportPath);
                if (!Directory.Exists(RunningEnv.ProgramParam.ReportDataPath))
                    Directory.CreateDirectory(RunningEnv.ProgramParam.ReportDataPath);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /*
         * 删除数据并无视异常
         */
        static public void DeleteOldFile()
        {
            try
            {
                DirectoryInfo floder = new DirectoryInfo(RunningEnv.ProgramParam.ToCheckTxtPaperPath);
                FileInfo[] fileInfo = floder.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();

                floder = new DirectoryInfo(RunningEnv.ProgramParam.ReportPath);
                fileInfo = floder.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();

                floder = new DirectoryInfo(RunningEnv.ProgramParam.ReportDataPath);
                fileInfo = floder.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();
            }
            catch
            { }
        }

        /*
         * 直接退出程序
         */
        static public void ForceStop()
        {
            Environment.Exit(0);
        }

        /*
         * 保存设置到配置文件，忽略错误
         */
        static public void SaveConfig(RunningEnv runningEnv)
        {
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                fs = new FileStream("config.ini", FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs);
                sw.WriteLine(runningEnv.CheckData.CheckThreshold);
                sw.WriteLine(runningEnv.SettingData.CheckThreadCnt);
                sw.WriteLine(runningEnv.SettingData.ConvertThreadCnt);
                if (runningEnv.SettingData.SuportPdf)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (runningEnv.SettingData.SuportTxt)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (runningEnv.SettingData.SuportDoc)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (runningEnv.SettingData.SuportDocx)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
            }
            catch
            { 
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        /*
         * 读取配置文件，忽略错误
         */
        static public RunningEnv ReadConfig(MainForm mainForm)
        {
            RunningEnv runningEnv = new RunningEnv(mainForm);
            if (File.Exists("config.ini"))
            {
                StreamReader sr = null;
                try
                {
                    sr = new StreamReader("config.ini", Encoding.GetEncoding("GBK"));
                    runningEnv.CheckData.CheckThreshold = int.Parse(sr.ReadLine());
                    runningEnv.SettingData.CheckThreadCnt = int.Parse(sr.ReadLine());
                    runningEnv.SettingData.ConvertThreadCnt = int.Parse(sr.ReadLine());
                    string line = sr.ReadLine();
                    if (line == "0")
                        runningEnv.SettingData.SuportPdf = false;
                    line = sr.ReadLine();
                    if (line == "0")
                        runningEnv.SettingData.SuportTxt = false;
                    line = sr.ReadLine();
                    if (line == "0")
                        runningEnv.SettingData.SuportDoc = false;
                    line = sr.ReadLine();
                    if (line == "0")
                        runningEnv.SettingData.SuportDocx = false;
                }
                catch
                {
                }
                finally
                {
                    if (sr != null) sr.Close();
                }
            }
            return runningEnv;
        }

        /*
         * 重置
         */
        static public void ResetProgram()
        {
            //确认提示
            DialogResult dr = MessageBox.Show("重置后系统将关闭，是否继续重置？", "提示", MessageBoxButtons.OKCancel);
            if (dr != DialogResult.OK)
                return;
            //删除历史文件
            DeleteOldFile();
            //强制退出程序
            ForceStop();
        }

        /*
         * 展示RunningEnv中的错误信息
         */
        static public void ShowErrorMsg(RunningEnv runningEnv)
        {
            if (runningEnv.CheckingData.ErrorPaperList.Count > 0)
            {
                StringBuilder error_msg = new StringBuilder("以下论文出现错误，请检查：\r\n");
                foreach (string paper_name in runningEnv.CheckingData.ErrorPaperList)
                {
                    error_msg.Append("●" + paper_name + "\r\n");
                }
                error_msg.Append("\r\n可能是以下原因导致：\r\n1、pdf或word文件存在只读密码、打开密码或防复制保护，请将保护取消。\r\n" +
                "2、pdf或word文件为图片或扫描件。系统无法对扫描件和图片进行查重，建议使用office word来生成pdf。\r\n" +
                "3、出错文件已经损坏无法正常打开，请尝试将上述文件诸葛打开检查。\r\n" +
                "4、文件的大小、字数小于所设阈值。\r\n" +
                "5、出错的word文件为非标文件，请使用office word打开，另存为新的word文件后再进行查重。");
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, error_msg.ToString(), "提示");
                }));
            }

        }

        /*
         * 计算并更新进度条
         */
        static public void UpdateCheckingProgress(RunningEnv runningEnv)
        {
            try
            {
                int num1 = Directory.GetFiles(runningEnv.CheckData.ToCheckPaperPath).Length;
                int num2 = Directory.GetFiles(RunningEnv.ProgramParam.ToCheckTxtPaperPath).Length;
                int num3 = Directory.GetFiles(RunningEnv.ProgramParam.ReportPath).Length;
                int num4 = Directory.GetFiles(runningEnv.CheckData.FinalReportPath).Length;
                if (num2 > 0)
                {
                    int temp = (int)(35.0 * Math.Min(1.0, 1.0 * num2 / Math.Max(1, num1))
                        + 60.0 * Math.Min(1.0, 1.0 * num3 / Math.Max(num2, 1))
                        + 5.0 * Math.Min(1.0, 1.0 * num4 / Math.Max(num3, 1)));
                    temp = Math.Min(100, temp + 1);
                    if (temp >= 0)
                    {
                        runningEnv.UIContext.BeginInvoke(new Action(() =>
                        {
                           runningEnv.UIContext.pbCheckingProgress.Value = temp;
                        }));
                    }
                }
            }
            catch (Exception)
            { }
        }

        static public void AdsMessage()
        {
            MessageBox.Show("本项目可以提供c#版本的相关技术支持，也可定制java版sdk开发包及技术支持，商业合作请联系微信/QQ：654062779。\r\n本提示您可自行修改源代码去除", "提示");
        }

    }
}
