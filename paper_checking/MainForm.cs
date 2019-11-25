using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.text;
using System.Runtime.InteropServices;

namespace paper_checking
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }
        
        string paper_source = "PaperSource/";
        string to_check_paper = "toCheckPaper/";
        const string txt_paper_source = "txtPaperSource/";
        const string to_check_txt_paper = "toCheckTxtPaper/";
        const string report = "report/";
        const string report_data = "report/Data/";
        const int min_bytes = 80; //格式转换后文件的最小大小。可根据情况自行修改。
        const int min_words = 10; //格式转换后文件内的最少字符数。可根据情况自行修改。
        const int max_words = 99998; //格式转换后文件内的最大字符数。不可修改。
        const string security_key = "Ubzrfax@3&Yl1rf&cw7ZE4zXsm8ZdIAtyJZ71L48f3yW*TXzylZq7Hqb1moG*xeQQnkFdkqYYXFfyPAS$CeETMw#1qDAPJehBM8";

        //DLL引用
        [DllImport(@"paper_check.dll", EntryPoint = "real_check", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        extern static int real_check(string s1, string s2, string s3, string s4, string s5, string s6, string s7, string s8);

        string Get_text_from_pdf_by_pdfbox(string path)
        {
            PDDocument pdffile = PDDocument.load(new java.io.File(path));
            PDFTextStripper pdfStripper = new PDFTextStripper();
            string text = pdfStripper.getText(pdffile);
            pdffile.close();
            //File.WriteAllText(dst, text, Encoding.GetEncoding("GBK"));
            return text;
        }

        string Get_text_from_word_by_spire(string path, Spire.Doc.Document doc)
        {
            doc.LoadFromFile(path);
            try
            {
                doc.Sections[0].HeadersFooters.Header.ChildObjects.Clear();
                doc.Sections[0].HeadersFooters.Footer.ChildObjects.Clear();
            }
            catch
            { }
            return doc.GetText();
        }

        /*
         * 去除特殊字符、目录和参考文献
         */
        void Txtfile_format(string text, string path)
        {
            if (path.ToLower().EndsWith(".doc.txt") || path.ToLower().EndsWith(".docx.txt") || path.ToLower().EndsWith(".txt.txt"))
            {
                text = text.Replace("#", "").Replace('\r', '#').Replace('\n', '#');
                text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！\#]", "");
                text = new Regex("[#]+").Replace(text, "@@").Trim();
            }
            else
            {
                text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！]", "");
                text = Regex.Replace(text, @"\s", string.Empty);
            }
            //text = Regex.Replace(text, @"\s", string.Empty);
            //str1 = Regex.Replace(Regex.Replace(str1, "(?<=[\u4e00-\u9fa5])\\s+(?=[\u4e00-\u9fa5])", string.Empty), "(?<=[a-z])\\s+(?=[a-z])", " ");

            if (text.Length < min_words)
            {
                File.WriteAllText(path, " ", Encoding.GetEncoding("GBK"));
                return;
            }

            if (1.0 * text.IndexOf("参考文献") / text.Length < 0.2)
                text = text.Substring(text.IndexOf("参考文献") + 4);
            if (text.IndexOf("引言") >= 0 && 1.0 * text.IndexOf("引言") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("引言") + 2);
            else if (text.IndexOf("绪论") >= 0 && 1.0 * text.IndexOf("绪论") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("绪论") + 2);
            else if (1.0 * text.IndexOf("序言") / text.Length < 0.1)
                text = text.Substring(text.IndexOf("序言") + 2);

            int cntckwx = 0;
            while (text.LastIndexOf("参考文献") > 0 && 1.0 * text.LastIndexOf("参考文献") / text.Length > 0.85 && cntckwx < 3)
            {
                text = text.Substring(0, text.LastIndexOf("参考文献"));
                cntckwx++;
            }

            text = text.Replace("（）", "").Replace("“”", "").Replace("，，", "，");

            if (text.Length > max_words)
            {
                text = text.Substring(0, max_words);
            }

            text = text.Trim("@".ToCharArray());
            File.WriteAllText(path, text, Encoding.GetEncoding("GBK"));
        }

        /*
         * 报错信息
         */
        private StringBuilder error_msg = new StringBuilder("以下论文出现错误，请检查：\r\n");
        void Show_convert_error(bool show)
        {
            if (error_msg.Length > 15)
            {
                if (show)
                {
                    error_msg.Append("\r\n可能是以下原因导致：\r\n1、pdf或word文件存在只读密码、打开密码或防复制保护，请将保护取消。\r\n" +
                                    "2、pdf或word文件为图片或扫描件。系统无法对扫描件和图片进行查重，建议使用office word来生成pdf。\r\n" +
                                    "3、出错文件已经损坏无法正常打开，请尝试将上述文件诸葛打开检查。\r\n" +
                                    "4、出错的word文件为非标文件，请使用office word打开，另存为新的word文件后再进行查重。");

                    MessageBox.Show(this, error_msg.ToString());
                }
                error_msg.Clear();
                error_msg.Append("以下论文出现错误，请检查：\r\n");
            }
        }

        /*
         * 创建存放论文和论文库的文件夹
         */
        void Folder_check_and_recover()
        {
            try
            {
                if (!Directory.Exists(paper_source))
                    Directory.CreateDirectory(paper_source);
                if (!Directory.Exists(txt_paper_source))
                    Directory.CreateDirectory(txt_paper_source);
                if (!Directory.Exists(to_check_paper))
                    Directory.CreateDirectory(to_check_paper);
                if (!Directory.Exists(to_check_txt_paper))
                    Directory.CreateDirectory(to_check_txt_paper);
                if (!Directory.Exists(report))
                    Directory.CreateDirectory(report);
                if (!Directory.Exists(report_data))
                    Directory.CreateDirectory(report_data);
            }
            catch (Exception)
            {
                MessageBox.Show("查重系统无法创建所需要的文件夹！无法启动！");
                Environment.Exit(1);
            }
        }

        public bool Is_unsign(string value)
        {
            return Regex.IsMatch(value, @"^\d*[.]?\d*$");
        }

        /*
         * 配置文件缺失
         */
        void Config_check_and_recover()
        {
            if (textBox7.Text == "" || !Is_unsign(textBox7.Text) || textBox7.Text.Length > 2)
                textBox7.Text = "12";
            if (textBox8.Text == "" || !Is_unsign(textBox8.Text) || textBox7.Text.Length > 2)
                textBox8.Text = "3";
            if (textBox9.Text == "" || !Is_unsign(textBox9.Text) || textBox7.Text.Length > 2)
                textBox9.Text = "2";
        }

        /*
         * 界面控制
         */
        public void Set_component_state(bool state)
        {
            textBox7.Enabled = state;
            groupBox1.Enabled = state;
            groupBox2.Enabled = state;
            button1.Enabled = state;
            button3.Enabled = state;
            button5.Enabled = state;
            button6.Enabled = state;
            button12.Enabled = state;
            button13.Enabled = state;
            checkBox1.Enabled = state;
            checkBox7.Enabled = state;
            comboBox1.Enabled = state;
        }

        /*
         * 删除数据并无视异常
         */
        public void Delete_check_data_file()
        {
            try
            {
                DirectoryInfo textFolder33 = new DirectoryInfo(to_check_txt_paper);
                DirectoryInfo textFolder11 = new DirectoryInfo(report);
                DirectoryInfo textFolder22 = new DirectoryInfo(report_data);
                FileInfo[] fileInfo = textFolder33.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();

                fileInfo = textFolder11.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();

                fileInfo = textFolder22.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                    NextFile.Delete();
            }
            catch
            {}
        }

        /*
         * 重置
         */
        public void Clear_env_data_file()
        {
            DialogResult dr = MessageBox.Show("重置后系统将关闭，是否继续重置？", "重置系统", MessageBoxButtons.OKCancel);
            if (dr != DialogResult.OK)
                return;

            Delete_check_data_file();
            Force_stop();
        }

        /*
         * 导出重率统计表
         */
        public void Export_table_report(string path)
        {
            DirectoryInfo data_floder = new DirectoryInfo(report_data);
            StringBuilder table_report = new StringBuilder("文件名, 重复率（%）\r\n");
            foreach (FileInfo file in data_floder.GetFiles())
            {
                StreamReader sr = new StreamReader(file.FullName, Encoding.GetEncoding("GBK"));
                table_report.Append(file.Name.Replace(",", "-") + ", " + sr.ReadLine() + "\r\n");
                if (sr != null) sr.Close();
            }
            File.WriteAllText(path + "\\重率统计表.csv", table_report.ToString(), Encoding.GetEncoding("GBK"));
        }

        public void Convert_PDF(object p_arr)
        {
            string[] param = (string[])p_arr;
            int taskNo = int.Parse(param[0].ToString());
            DirectoryInfo sourceFolder = new DirectoryInfo(param[1].ToString());
            DirectoryInfo textFolder = new DirectoryInfo(param[2].ToString());

            if (!sourceFolder.Exists || !textFolder.Exists)
            {
                MessageBox.Show(this, "文件夹丢失，任务失败！");
                return;
            }

            FileInfo[] fileInfo = sourceFolder.GetFiles();
            Spire.Doc.Document doc = new Spire.Doc.Document();

            int addNumo = int.Parse(textBox9.Text);
            for (int FileInfoNo = taskNo; FileInfoNo < fileInfo.Length; FileInfoNo += addNumo)
            {
                FileInfo NextFile = fileInfo[FileInfoNo];

                string path = sourceFolder.FullName + "\\" + NextFile.Name;
                string t = textFolder.FullName + "\\" + NextFile.Name + ".txt";

                if (File.Exists(t))
                {
                    continue;
                }

                try
                {
                    if (path.ToLower().EndsWith(".pdf") && checkBox2.Checked)
                    {
                        string text = Get_text_from_pdf_by_pdfbox(path);
                        Txtfile_format(text, t);
                    }
                    else if (path.ToLower().EndsWith(".txt") && checkBox5.Checked)
                    {
                        string text = File.ReadAllText(path, Encoding.GetEncoding("GBK"));
                        Txtfile_format(text, t);
                    }
                    else if(path.ToLower().EndsWith(".doc") && checkBox3.Checked)
                    {
                        string text = Get_text_from_word_by_spire(path, doc);
                        Txtfile_format(text, t);
                    }
                    else if (path.ToLower().EndsWith(".docx") && checkBox4.Checked)
                    {
                        string text = Get_text_from_word_by_spire(path, doc);
                        Txtfile_format(text, t);
                    }
                    else
                    {
                        continue;
                    }
                    if (new FileInfo(t).Length <= min_bytes)
                    {
                        if (File.Exists(t))
                        {
                            File.Delete(t);
                        }
                        throw new Exception();
                    }
                }
                catch(Exception e)
                {
                    error_msg.Append("●" + NextFile.Name + ";\r\n");
                }

            }
        }

        void Start_convert_PDF(object p_arr)
        {
            button4.Enabled = false;
            button2.Enabled = false;
            string[] part_param = (string[])p_arr;
            List<Thread> temp = new List<Thread>();
            for (int i = 0; i < int.Parse(part_param[0]); i++)
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(Convert_PDF));
                string[] param = new string[3] { i.ToString(), part_param[1], part_param[2] };
                t1.Start(param);
                temp.Add(t1);
            }
            foreach (Thread item in temp)
            {
                item.Join();
            }
            temp.Clear();
            if (part_param[3].Equals("1"))
            {
                Show_convert_error(true);
            }
            button4.Enabled = true;
            button2.Enabled = true;
        }

        private void Button_add_to_paper_library(object sender, EventArgs e)
        {
            if (button3.Enabled == false || button4.Enabled == false || int.Parse(textBox9.Text) <= 0)
            {
                MessageBox.Show(this, "请等待当前任务完成后再进行论文库更新！");
                return;
            }
            Thread t1 = new Thread(new ParameterizedThreadStart(Start_convert_PDF));
            string[] param = new string[4] { textBox9.Text, paper_source, txt_paper_source, "1" };
            t1.Start(param);
        }

        //查重核心dll的多线程包装外壳
        void real_check_shell(object p_arr)
        {
            string[] part_param = (string[])p_arr;
            int res = real_check(part_param[0], part_param[1], part_param[2], part_param[3], part_param[4], part_param[5], part_param[6], part_param[7]);
            Exep_exited(res);
        }

        public void Check_paper_thread()
        {
            Set_component_state(false);
            if (checkBox1.Checked == false)
            {
                Delete_check_data_file();
            }

            Thread t2 = new Thread(new ParameterizedThreadStart(Start_convert_PDF));
            string[] param = new string[4] { textBox9.Text, to_check_paper, to_check_txt_paper, "0" };
            t2.Start(param);
            t2.Join();

            for (int i = 0; i < int.Parse(textBox8.Text); i++)
            {
                try
                {
                    string[] real_check_param;
                    if (comboBox1.SelectedIndex == 1)
                    {
                        real_check_param = new string[8] { textBox7.Text, textBox8.Text, i.ToString(), security_key, to_check_txt_paper, to_check_txt_paper, report, report_data };
                    }
                    else
                    {
                        real_check_param = new string[8] { textBox7.Text, textBox8.Text, i.ToString(), security_key, to_check_txt_paper, txt_paper_source, report, report_data };
                    }
                    Thread real_check_t = new Thread(new ParameterizedThreadStart(real_check_shell));
                    real_check_t.Start(real_check_param);
                }
                catch (Exception e)
                {
                    MessageBox.Show(this, i.ToString() + "号线程执行失败，请检查xc_core文件是否存在！");
                }
            }

        }

        string ReportSavePath = "";
        Thread t111;
        private void Button_start_checking(object sender, EventArgs e)
        {
            if (int.Parse(textBox7.Text) <= 0 || int.Parse(textBox8.Text) <= 0 || int.Parse(textBox9.Text) <= 0)
            {
                MessageBox.Show(this, "线程数和查重阈值必须都大于0！");
                return;
            }

            if (button2.Enabled == false || button3.Enabled == false || button4.Enabled == false)
            {
                MessageBox.Show(this, "请等待库更新完毕再进行查重！");
                return;
            }
            
            if (textBox1.Text == "")
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ReportSavePath = dialog.SelectedPath;
                    textBox1.Text = ReportSavePath;
                }
                else
                {
                    return;
                }
            }
            else
            {
                ReportSavePath = textBox1.Text;
            }
            B_message();
            t111 = new Thread(new ThreadStart(Check_paper_thread));
            t111.Start();
        }

        private static Semaphore sem = new Semaphore(1, 1);
        int curTSUM = 0;
        void Exep_exited(int res)
        {
            sem.WaitOne();
            curTSUM++;
            if (res != 0)
            {
                MessageBox.Show(this, "存在未知的查重失败，请排查：待查论文是否有错误的格式？查重系统key文件是否存在？");
            }
            if (curTSUM >= int.Parse(textBox8.Text))
            {
                curTSUM = 0;
                if (ReportSavePath != "")
                {
                    Reporter(ReportSavePath);
                }
                else
                {
                    MessageBox.Show(this, "未选择文件夹，无法导出查重报告！");
                }
                Show_convert_error(true);
                Set_component_state(true);
                GC.Collect();
            }
            sem.Release();
        }

        private void Button_look_report(object sender, EventArgs e)
        {
            try
            {
                ReportListForm f = new ReportListForm(to_check_txt_paper, report, report_data);
                f.Show();
            }
            catch { }
        }

        private void Reporter(string path)
        {
            string exc = "以下论文的报告导出失败，请检查：";

            DirectoryInfo sourceFolder = new DirectoryInfo(report);
            if (!sourceFolder.Exists)
            {
                MessageBox.Show(this, "report文件夹丢失！导出失败！");
                return;
            }
            FileInfo[] fileInfo = sourceFolder.GetFiles();
            for (int i = 0; i < fileInfo.Length; i++)
            {
                try
                {
                    ReportDetailForm f = new ReportDetailForm(fileInfo[i].Name.Substring(0, fileInfo[i].Name.Length - 4), to_check_txt_paper, report, report_data);
                    f.richTextBox1.SaveFile(path + "\\" + fileInfo[i].Name.Substring(0, fileInfo[i].Name.Length - 4) + ".rtf");
                }
                catch
                {
                    exc += fileInfo[i].Name + "、";
                }
            }
            if (checkBox7.Checked) {
                Export_table_report(path);
            }
            TopMost = true;
            MessageBox.Show(this, "导出完成！");
            if (exc != "以下论文的报告导出失败，请检查：")
            {
                MessageBox.Show(this, exc);
            }
            TopMost = false;
            ReportSavePath = "";
        }

        private void Button_export_report(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Reporter(dialog.SelectedPath);
            }
        }

        /*
         * 阻止数字以外的字符输入
         */
        private void TextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && !char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        /*
         * 退出时写入配置文件
         */
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (tabControl1.TabPages.Count <= 1)
                return;

            try
            {
                FileStream fs;
                StreamWriter sw;
                fs = new FileStream("config.ini", FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs);
                sw.WriteLine(textBox7.Text);
                sw.WriteLine(textBox8.Text);
                sw.WriteLine(textBox9.Text);
                if (checkBox2.Checked)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (checkBox5.Checked)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (checkBox3.Checked)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                if (checkBox4.Checked)
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");
                sw.Close();
                fs.Close();
            }
            catch
            {}

            if (button3.Enabled == false || button2.Enabled == false)
            {
                Force_stop();
            }
        }

        /*
         * 启动时读取配置文件
         */
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            if (File.Exists("config.ini"))
            {
                StreamReader sr = new StreamReader("config.ini", Encoding.GetEncoding("GBK"));
                textBox7.Text = sr.ReadLine();
                textBox8.Text = sr.ReadLine();
                textBox9.Text = sr.ReadLine();

                string line = sr.ReadLine();
                if (line == "0")
                    checkBox2.Checked = false;

                line = sr.ReadLine();
                if (line == "0")
                    checkBox5.Checked = false;

                line = sr.ReadLine();
                if (line == "0")
                    checkBox3.Checked = false;

                line = sr.ReadLine();
                if (line == "0")
                    checkBox4.Checked = false;

                if (sr != null) sr.Close();
            }
            else
            {
                Button9_Click(sender, e);
            }
            Config_check_and_recover();
            Folder_check_and_recover();
        }

        private void Force_stop()
        {
            if (button3.Enabled == false || button2.Enabled == false)
            {
                DialogResult dr = MessageBox.Show("有任务正在进行，强制退出有可能丢失数据，是否继续本次强制退出？。", "强制退出", MessageBoxButtons.OKCancel);
                if (dr != DialogResult.OK)
                {
                    return;
                }
            }
            try
            {
                Process[] myproc = Process.GetProcesses();
                foreach (Process item in myproc)
                {
                    if (item.ProcessName == "xc_core")
                    {
                        item.Kill();
                    }
                }
            }
            catch { }
            try
            {
                t111.Abort();
            }
            catch { }
            Environment.Exit(0);
        }

        private void Button_select_path2(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
            }
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            int processorCount = Environment.ProcessorCount - 2;
            if (processorCount > 1)
            {
                textBox8.Text = processorCount.ToString();
                textBox9.Text = processorCount.ToString();
            }
            else
            {
                textBox8.Text = "1";
                textBox9.Text = "1";
            }
        }

        private void Button_reset_system(object sender, EventArgs e)
        {
            Clear_env_data_file();
        }

        private void button_force_stop(object sender, EventArgs e)
        {
            Force_stop();
        }

        private void Button_select_path1(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = dialog.SelectedPath;
                to_check_paper = textBox5.Text + "\\";
            }
        }

        private void Button_select_path3(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox6.Text = dialog.SelectedPath;
                paper_source = textBox6.Text + "\\";
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            new Licence().ShowDialog();
        }

        private void B_message()
        {
            string[] cfiles = Directory.GetFiles(textBox5.Text); 
            string[] sfiles = Directory.GetFiles(txt_paper_source);
            if (int.Parse(textBox8.Text) > 2 || int.Parse(textBox9.Text) > 2 || cfiles.Length > 10 || sfiles.Length > 100)
            {
                MessageBox.Show("本项目可以提供c#版本的相关技术支持，也可提供java版sdk开发包及技术支持，商业合作请联系微信/QQ：654062779。\r\n当线程数大于2/待查论文数超过10/论文库中论文超过100时，将会显示此提示，您可自行修改源代码去除");
            }
        }
    }
   
}
