using paper_checking.PaperCheck.Convert;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace paper_checking.PaperCheck
{
    class PaperManager
    {
        private readonly RunningEnv runningEnv;

        public PaperManager(RunningEnv env)
        {
            runningEnv = env;
        }

        //DLL引用
        [DllImport(@"paper_check.dll", EntryPoint = "real_check", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = false, CallingConvention = CallingConvention.StdCall)]
        extern static int real_check(string s1, string s2, string s3, string s4, string s5, string s6, string s7, string s8);

        /*
         * 将文件转换为文本
         */
        private void FileConvert(object thread_payload)
        {
            //获取参数
            string[] param = (string[])thread_payload;
            int taskNo = int.Parse(param[0].ToString());
            DirectoryInfo sourceFolder = new DirectoryInfo(param[1]);
            DirectoryInfo textFolder = new DirectoryInfo(param[2]);
            
            //判断源文件夹和目标文件夹是否存在
            if (!sourceFolder.Exists || !textFolder.Exists)
            {
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, "数据损坏，任务失败！", "错误");
                }));
                return;
            }

            //获取源文件夹中的待转换文件
            FileInfo[] fileInfo = Utils.GetFileInfoRecursion(sourceFolder).ToArray();

            //按本线程所分配的任务进行转换
            int addNumo = int.Parse(runningEnv.SettingData.ConvertThreadCnt.ToString());
            for (int FileInfoNo = taskNo; FileInfoNo < fileInfo.Length; FileInfoNo += addNumo)
            {
                FileInfo NextFile = fileInfo[FileInfoNo];

                string path = NextFile.FullName;
                string real_dis_file_name = Regex.Replace(NextFile.Name, @"[^\u4e00-\u9fa5\u0022\《\》\（\）\—\；\，\。\“\”\！\#\\_\-\.\,\:\(\)\'\[\]\【\】\+\·\：\<\>\w]", string.Empty);
                foreach (char rInvalidChar in Path.GetInvalidFileNameChars())
                {
                    real_dis_file_name = real_dis_file_name.Replace(rInvalidChar.ToString(), string.Empty);
                }
                string dist_path = textFolder.FullName + Path.DirectorySeparatorChar + real_dis_file_name + ".txt";

                //文件已经被转换则忽略该文件
                if (File.Exists(dist_path))
                {
                    continue;
                }

                ConverterFactory converterFactory = new ConverterFactory();

                try
                {
                    //获取文件后缀
                    string file_type = path.ToLower().Split('.')[Math.Max(path.ToLower().Split('.').Length - 1, 0)];
                    //获取转换器
                    ConvertCore file_converter = converterFactory.GetConverter(file_type, runningEnv);
                    string text = "";
                    if (file_converter != null)//如果可以获得到转换器
                    {
                        //获取文本
                        text = file_converter.ConvertToString(path, runningEnv.CheckData.Blocklist);
                        if (text != null && text.Length > 0)
                        {
                            if (text.Length > RunningEnv.ProgramParam.MaxWords)
                            {
                                //舍弃过长的部分
                                text = text.Substring(0, RunningEnv.ProgramParam.MaxWords);
                            }
                            //写入目标路径
                            File.WriteAllText(dist_path, text, Encoding.GetEncoding("GBK"));
                        }
                    }
                    else
                    {
                        //没有获取到转换器则忽略
                        continue;
                    }
                    
                    //如果转换后的文件不符合所设阈值则删除
                    if (new FileInfo(dist_path).Length <= runningEnv.CheckData.MinBytes || text.Length <= runningEnv.CheckData.MinWords)
                    {
                        if (File.Exists(dist_path))
                        {
                            File.Delete(dist_path);
                        }
                        //并抛出一个异常
                        throw new Exception();
                    }
                }
                catch (Exception e)
                {
                    runningEnv.CheckingData.ErrorPaperList.AddLast(NextFile.Name);
                }
            }
        }

        /*
         * 独立调用文件转换
         */
        public void StartFileConvertStandalone()
        {
            //判断是否已经有一个可能会进行文件转换的任务
            if (!Monitor.TryEnter(RunningEnv.EnvRunningLock))
            {
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, "不可以同时执行两个任务！", "提示");
                }));
                return;
            }
            //UI更新
            MainForm.SetPartComponentState(runningEnv.UIContext, false);
            //启动转换任务
            string[] param = new string[2] { runningEnv.LibraryData.PaperSourcePath, RunningEnv.ProgramParam.TxtPaperSourcePath };
            StartFileConvert(param);
            //UI更新
            MainForm.SetPartComponentState(runningEnv.UIContext, true);
            //锁释放
            Monitor.Exit(RunningEnv.EnvRunningLock);
        }

        /*
         * 过程中调用文件转换
         */
        private void StartFileConvert(object thread_payload)
        {
            string[] param = (string[])thread_payload;
            //线程集合
            List<Thread> temp = new List<Thread>();
            for (int i = 0; i < runningEnv.SettingData.ConvertThreadCnt; i++)
            {
                //启动转换线程
                Thread t1 = new Thread(new ParameterizedThreadStart(FileConvert));
                string[] new_param = new string[3] { i.ToString(), param[0], param[1] };
                t1.Start(new_param);
                //添加到线程集合
                temp.Add(t1);
            }
            //等待转换完毕
            foreach (Thread item in temp)
            {
                item.Join();
            }
            temp.Clear();
        }

        /*
         * 调用动态链接库方法的包装方法
         */
        private void RealCheckShell(object thread_payload)
        {
            //获取参数
            string[] part_param = (string[])thread_payload;
            int res = real_check(part_param[0], part_param[1], part_param[2], part_param[3], part_param[4], part_param[5], part_param[6], part_param[7]);
            //如果动态链接库异常返回值
            string errMsg = "";
            //记录错误信息
            switch (res)
            {
                case 0:
                    return;
                case 9:
                    errMsg = "key文件读取失败。建议排查：1、key文件是否存在；2、是否拥有key文件的读取权限。";
                    break;
                default:
                    errMsg = "出现未知的错误。建议排查：1、比对库中文件的文件名是否包含特殊字符；2、操作系统寻址位是否可paper_check.dll一致；3、查重系统所在文件夹路径是否存在特殊字符。";
                    break;
            }
            //错误计数加1
            Interlocked.Increment(ref runningEnv.CheckingData.DllErrCount);
            //判断是否是第一次错误
            if (!errMsg.Equals("") && runningEnv.CheckingData.DllErrCount == 1)
            {
                //显示错误信息
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, errMsg, "错误");
                }));
            }
        }

        /*
         * 更新主界面进度条的线程
         */
        private void StartProgressBar()
        {
            while (true)
            {
                //如果查重任务结束则退出
                if (runningEnv.CheckingData.IsFinisth)
                {
                    //进度条恢复为0
                    runningEnv.UIContext.BeginInvoke(new Action(() =>
                    {
                        runningEnv.UIContext.pbCheckingProgress.Value = 0;
                    }));
                    break;
                }
                Utils.UpdateCheckingProgress(runningEnv);
                Thread.Sleep(1500);
            }
        }

        /*
         * 开始查重
         */
        public void StartCheckPaper()
        {
            Utils.AdsMessage();
            //判断是否已经有一个正在进行的任务
            if (!Monitor.TryEnter(RunningEnv.EnvRunningLock))
            {
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, "不可以同时执行两个任务！", "提示");
                }));
                return;
            }

            //更新UI
            MainForm.SetComponentState(runningEnv.UIContext, false);

            //如果不是中断恢复则删除历史文件
            if (!runningEnv.CheckData.Recover)
            {
                Utils.DeleteOldFile();
            }

            //启动更新进度条的线程
            Thread progressBarThread = new Thread(new ThreadStart(StartProgressBar));
            progressBarThread.Start();

            //启动文件转换任务
            Thread fileConvertThread = new Thread(new ParameterizedThreadStart(StartFileConvert));
            string[] param = new string[2] { runningEnv.CheckData.ToCheckPaperPath, RunningEnv.ProgramParam.ToCheckTxtPaperPath };
            fileConvertThread.Start(param);
            fileConvertThread.Join();

            //启动查重任务
            LinkedList<Thread> realCheckThreadList = new LinkedList<Thread>();
            for (int i = 0; i < runningEnv.SettingData.CheckThreadCnt; i++)
            {
                try
                {
                    string[] real_check_param;
                    if (runningEnv.CheckData.CheckWay == 1)
                    {
                        //如果是横向查重
                        real_check_param = new string[8] { runningEnv.CheckData.CheckThreshold.ToString(),
                                                            runningEnv.SettingData.CheckThreadCnt.ToString(), 
                                                            i.ToString(),
                                                            RunningEnv.ProgramParam.SecurityKey,
                                                            RunningEnv.ProgramParam.ToCheckTxtPaperPath,
                                                            RunningEnv.ProgramParam.ToCheckTxtPaperPath,
                                                            RunningEnv.ProgramParam.ReportPath,
                                                            RunningEnv.ProgramParam.ReportDataPath };
                    }
                    else
                    {
                        //如果是纵向查重
                        real_check_param = new string[8] { runningEnv.CheckData.CheckThreshold.ToString(),
                                                            runningEnv.SettingData.CheckThreadCnt.ToString(),
                                                            i.ToString(),
                                                            RunningEnv.ProgramParam.SecurityKey,
                                                            RunningEnv.ProgramParam.ToCheckTxtPaperPath,
                                                            RunningEnv.ProgramParam.TxtPaperSourcePath,
                                                            RunningEnv.ProgramParam.ReportPath,
                                                            RunningEnv.ProgramParam.ReportDataPath };
                    }
                    //启动对应线程
                    Thread real_check_t = new Thread(new ParameterizedThreadStart(RealCheckShell));
                    realCheckThreadList.AddLast(real_check_t);
                    real_check_t.Start(real_check_param);
                }
                catch (Exception e)
                {
                    //如果遇到未知异常，错误计数加1
                    Interlocked.Increment(ref runningEnv.CheckingData.ThreadErrCount);
                    //判断是否是第一次错误
                    if (runningEnv.CheckingData.ThreadErrCount == 1)
                    {
                        //显示错误信息
                        runningEnv.UIContext.BeginInvoke(new Action(() =>
                        {
                            MessageBox.Show("未知线程错误，全部或部分文件查重失败。请在查重结束后检查！", "错误");
                        }));
                    }
                }
            }

            //等待查重结束
            foreach (Thread thread in realCheckThreadList)
            {
                thread.Join();
            }

            //如果有可导出的查重报告，则导出查重报告
            if (Directory.GetFiles(RunningEnv.ProgramParam.ReportPath).Length > 0 && Directory.GetFiles(RunningEnv.ProgramParam.ReportDataPath).Length > 0)
            {
                ExportReport();
            }

            //内存垃圾回收
            GC.Collect();
            //UI更新
            MainForm.SetComponentState(runningEnv.UIContext, true);
            //设置查重结束标记
            runningEnv.CheckingData.IsFinisth = true;
            //锁释放
            Monitor.Exit(RunningEnv.EnvRunningLock);
            //显示报错信息（如有）
            Utils.ShowErrorMsg(runningEnv);
        }

        /*
         * 导出查重报告
         */
        public void ExportReport()
        {
            //检查所需数据是否存在
            if (!Directory.Exists(RunningEnv.ProgramParam.ReportPath)||
                !Directory.Exists(RunningEnv.ProgramParam.ReportDataPath) ||
                !Directory.Exists(RunningEnv.ProgramParam.ToCheckTxtPaperPath))
            {
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, "数据损坏，无法导出查重报告！", "错误");
                }));
                return;
            }

            //报错信息（部分数据丢失时报错）
            StringBuilder error_msg = new StringBuilder();
            DirectoryInfo sourceFolder = new DirectoryInfo(RunningEnv.ProgramParam.ReportPath);
            FileInfo[] fileInfo = sourceFolder.GetFiles();
            for (int i = 0; i < fileInfo.Length; i++)
            {
                try
                {
                    //逐一生成查重报告
                    ReportDetailForm f = new ReportDetailForm(fileInfo[i].Name.Substring(0, fileInfo[i].Name.Length - 4),
                                                                RunningEnv.ProgramParam.ToCheckTxtPaperPath,
                                                                RunningEnv.ProgramParam.ReportPath,
                                                                RunningEnv.ProgramParam.ReportDataPath);
                    f.richTextBox1.SaveFile(runningEnv.CheckData.FinalReportPath + Path.DirectorySeparatorChar + fileInfo[i].Name.Substring(0, fileInfo[i].Name.Length - 4) + ".rtf");
                    f.Dispose();
                }
                catch
                {
                    //如有异常则记录报错信息
                    error_msg.Append(fileInfo[i].Name + "、");
                }
            }
            //如果需要生成统计表
            if (runningEnv.CheckData.StatisTable)
            {
                ExportStatisTable();//生成统计表
            }
            //完成提示
            runningEnv.UIContext.BeginInvoke(new Action(() =>
            {
                MessageBox.Show(runningEnv.UIContext, "导出完成！");
            }));
            //显示报错信息（如有）
            if (error_msg.Length > 0)
            {
                runningEnv.UIContext.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(runningEnv.UIContext, "以下文件的报告导出失败，请检查：" + error_msg.ToString(), "提示");
                }));
            }
        }

        /*
         * 导出统计表
         */
        public void ExportStatisTable()
        {
            DirectoryInfo data_floder = new DirectoryInfo(RunningEnv.ProgramParam.ReportDataPath);
            StringBuilder table_report = new StringBuilder("文件名, 重复率（%）").Append(Environment.NewLine);
            foreach (FileInfo file in data_floder.GetFiles())
            {
                StreamReader sr = new StreamReader(file.FullName, Encoding.GetEncoding("GBK"));
                table_report.Append(file.Name.Replace(",", "-")).Append(", ").Append(sr.ReadLine().Replace(",", " ")).Append(Environment.NewLine);
                if (sr != null) sr.Close();
            }
            File.WriteAllText(runningEnv.CheckData.FinalReportPath + Path.DirectorySeparatorChar + "result.csv", table_report.ToString(), Encoding.GetEncoding("GBK"));
        }

    }
}
