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
        const int min_bytes = 18000;
        const int min_words = 100;
        const int max_words = 99998;
        const string security_key = "Ubzrfax@3&Yl1rf&cw7ZE4zXsm8ZdIAtyJZ71L48f3yW*TXzylZq7Hqb1moG*xeQQnkFdkqYYXFfyPAS$CeETMw#1qDAPJehBM8";

        string get_text_from_pdf_by_pdfbox(string path)
        {
            PDDocument pdffile = PDDocument.load(new java.io.File(path));
            PDFTextStripper pdfStripper = new PDFTextStripper();
            string text = pdfStripper.getText(pdffile);
            pdffile.close();
            //File.WriteAllText(dst, text, Encoding.GetEncoding("GBK"));
            return text;
        }
        string get_text_from_word_by_spire(string path, Spire.Doc.Document doc)
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

        void txtfile_format(string text, string path)
        {
            text = Regex.Replace(text, @"[^\u4e00-\u9fa5\《\》\（\）\——\；\，\。\“\”\！]", "");
            text = Regex.Replace(text, @"\s", string.Empty);
            //str1 = Regex.Replace(Regex.Replace(str1, "(?<=[\u4e00-\u9fa5])\\s+(?=[\u4e00-\u9fa5])", string.Empty), "(?<=[a-z])\\s+(?=[a-z])", " ");

            if(text.Length < min_words)
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

            if (text.LastIndexOf("参考文献") > 0 && 1.0 * text.LastIndexOf("参考文献") / text.Length > 0.8)
                text = text.Substring(0, text.LastIndexOf("参考文献"));

            text = text.Replace("（）", "").Replace("“”", "").Replace("，，", "，");

            if (text.Length > max_words)
            {
                text = text.Substring(0, max_words);
            }

            File.WriteAllText(path, text, Encoding.GetEncoding("GBK"));
        }

        StringBuilder error_msg = new StringBuilder("以下论文出现错误，请检查：\r\n");
        void show_convert_error(bool show)
        {
            if (error_msg.Length > 15)
            {
                error_msg.Append("\r\n可能是以下原因导致：\r\n1、pdf或word文件存在只读密码或打开密码，请将密码取消。\r\n2、pdf或word文件为图片或扫描件。系统无法对扫描件和图片进行查重，建议使用office word来生成pdf。\r\n3、上述出错文件无法正常打开，请尝试打开文件检查。");
                if (show)
                {
                    MessageBox.Show(this, error_msg.ToString());
                }
                error_msg.Clear();
                error_msg.Append("以下论文出现错误，请检查：\r\n");
            }
        }

        void folder_check_and_recover()
        {
            try
            {
                if (!Directory.Exists(paper_source))
                {
                    Directory.CreateDirectory(paper_source);
                }
                if (!Directory.Exists(txt_paper_source))
                {
                    Directory.CreateDirectory(txt_paper_source);
                }
                if (!Directory.Exists(to_check_paper))
                {
                    Directory.CreateDirectory(to_check_paper);
                }
                if (!Directory.Exists(to_check_txt_paper))
                {
                    Directory.CreateDirectory(to_check_txt_paper);
                }
                if (!Directory.Exists(report))
                {
                    Directory.CreateDirectory(report);
                }
                if (!Directory.Exists(report_data))
                {
                    Directory.CreateDirectory(report_data);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("查重系统无法创建所需要的文件夹！无法启动！");
                Environment.Exit(1);
            }
        }
        public bool is_unsign(string value)
        {
            return Regex.IsMatch(value, @"^\d*[.]?\d*$");
        }

        void config_check_and_recover()
        {
            if (textBox7.Text == "" || !is_unsign(textBox7.Text) || textBox7.Text.Length > 2)
                textBox7.Text = "12";
            if (textBox8.Text == "" || !is_unsign(textBox8.Text) || textBox7.Text.Length > 2)
                textBox8.Text = "3";
            if (textBox9.Text == "" || !is_unsign(textBox9.Text) || textBox7.Text.Length > 2)
                textBox9.Text = "2";
        }

        public void set_component_state(bool state)
        {
            textBox7.Enabled = state;
            groupBox1.Enabled = state;
            groupBox2.Enabled = state;
            button3.Enabled = state;
            button5.Enabled = state;
            button6.Enabled = state;
            button12.Enabled = state;
            checkBox1.Enabled = state;
            checkBox7.Enabled = state;
            comboBox1.Enabled = state;
        }

        public void delete_check_data_file()
        {
            try
            {
                DirectoryInfo textFolder33 = new DirectoryInfo(to_check_txt_paper);
                DirectoryInfo textFolder11 = new DirectoryInfo(report);
                DirectoryInfo textFolder22 = new DirectoryInfo(report_data);
                FileInfo[] fileInfo = textFolder33.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                {
                    NextFile.Delete();
                }
                fileInfo = textFolder11.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                {
                    NextFile.Delete();
                }
                fileInfo = textFolder22.GetFiles();
                foreach (FileInfo NextFile in fileInfo)
                {
                    NextFile.Delete();
                }
            }
            catch
            {}
        }

        public void clear_env_data_file()
        {
            DialogResult dr = MessageBox.Show("重置后系统将关闭，是否继续重置？", "重置系统", MessageBoxButtons.OKCancel);
            if (dr != DialogResult.OK)
            {
                return;
            }
            delete_check_data_file();
            force_stop();
        }

        public void export_table_report(string path)
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

        public void convert_PDF(object p_arr)
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
                String t = textFolder.FullName + "\\" + NextFile.Name + ".txt";

                if (File.Exists(t))
                {
                    continue;
                }

                try
                {
                    if (path.ToLower().EndsWith(".pdf") && checkBox2.Checked)
                    {
                        string text = get_text_from_pdf_by_pdfbox(path);
                        txtfile_format(text, t);
                    }
                    else if (path.ToLower().EndsWith(".txt") && checkBox5.Checked)
                    {
                        string text = File.ReadAllText(path, Encoding.GetEncoding("GBK"));
                        txtfile_format(text, t);
                    }
                    else if(path.ToLower().EndsWith(".doc") && checkBox3.Checked)
                    {
                        string text = get_text_from_word_by_spire(path, doc);
                        txtfile_format(text, t);
                    }
                    else if (path.ToLower().EndsWith(".docx") && checkBox4.Checked)
                    {
                        string text = get_text_from_word_by_spire(path, doc);
                        txtfile_format(text, t);
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

        void start_convert_PDF(object p_arr)
        {
            button4.Enabled = false;
            button2.Enabled = false;
            string[] part_param = (string[])p_arr;
            List<Thread> temp = new List<Thread>();
            for (int i = 0; i < int.Parse(part_param[0]); i++)
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(convert_PDF));
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
                show_convert_error(true);
            }
            button4.Enabled = true;
            button2.Enabled = true;
        }

        private void button_add_to_paper_library(object sender, EventArgs e)
        {
            if (button3.Enabled == false || button4.Enabled == false || int.Parse(textBox9.Text) <= 0)
            {
                MessageBox.Show(this, "请等待当前任务完成后再进行论文库更新！");
                return;
            }
            Thread t1 = new Thread(new ParameterizedThreadStart(start_convert_PDF));
            string[] param = new string[4] { textBox9.Text, paper_source, txt_paper_source, "1" };
            t1.Start(param);
        }

        public void check_paper_thread()
        {
            set_component_state(false);
            if (checkBox1.Checked == false)
            {
                delete_check_data_file();
            }

            Thread t1 = new Thread(new ParameterizedThreadStart(start_convert_PDF));
            string[] param = new string[4] { textBox9.Text, paper_source, txt_paper_source, "0" };
            t1.Start(param);
            t1.Join();
            
            param = new string[4] { textBox9.Text, to_check_paper, to_check_txt_paper, "0" };
            t1.Start(param);
            t1.Join();

            for (int i = 0; i < int.Parse(textBox8.Text); i++)
            {
                try
                {
                    Process myprocess = new Process();
                    string start_param;
                    if (comboBox1.SelectedIndex == 1)
                    {
                        start_param = (textBox7.Text + " " + textBox8.Text + " " + i.ToString() + " " + security_key + " "
                            + to_check_txt_paper + " " + to_check_txt_paper + " " + report + " " + report_data).Trim();
                    }
                    else
                    {
                        start_param = (textBox7.Text + " " + textBox8.Text + " " + i.ToString() + " " + security_key + " "
                            + to_check_txt_paper + " " + txt_paper_source + " " + report + " " + report_data).Trim();
                    }
                    ProcessStartInfo startInfo = new ProcessStartInfo("xc_core.exe", start_param);
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.CreateNoWindow = true;
                    startInfo.UseShellExecute = false;
                    myprocess.StartInfo = startInfo;
                    myprocess.EnableRaisingEvents = true;
                    myprocess.Exited += new EventHandler(exep_exited);
                    myprocess.Start();
                }
                catch
                {
                    MessageBox.Show(this, i.ToString() + "号线程执行失败，请检查xc_core文件是否存在！");
                }
            }
        }

        string ReportSavePath = "";
        Thread t111;
        private void button_start_checking(object sender, EventArgs e)
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
            t111 = new Thread(new ThreadStart(check_paper_thread));
            t111.Start();
        }

        static Semaphore sem = new Semaphore(1, 1);
        int curTSUM = 0;
        void exep_exited(object sender, EventArgs e)
        {
            sem.WaitOne();
            curTSUM++;
            Process process = (Process)sender;
            if (process.ExitCode != 0)
            {
                MessageBox.Show(this, "存在未知的查重失败，请排查：查重系统key文件是否存在？待查论文是否有错误的格式？");
            }
            if (curTSUM >= int.Parse(textBox8.Text))
            {
                curTSUM = 0;
                if (ReportSavePath != "")
                {
                    reporter(ReportSavePath);
                }
                else
                {
                    MessageBox.Show(this, "未选择文件夹，无法导出查重报告！");
                }
                show_convert_error(true);
                set_component_state(true);
            }
            sem.Release();
        }

        private void button_look_report(object sender, EventArgs e)
        {
            try
            {
                ReportListForm f = new ReportListForm(to_check_txt_paper, report, report_data);
                f.Show();
            }
            catch { }
        }

        private void reporter(string path)
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
                export_table_report(path);
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

        private void button_export_report(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                reporter(dialog.SelectedPath);
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && !Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (tabControl1.TabPages.Count <= 1)
            {
                return;
            }
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
                {
                    sw.WriteLine("1");
                }
                else
                {
                    sw.WriteLine("0");
                }
                if (checkBox5.Checked)
                {
                    sw.WriteLine("1");
                }
                else
                {
                    sw.WriteLine("0");
                }
                if (checkBox3.Checked)
                {
                    sw.WriteLine("1");
                }
                else
                {
                    sw.WriteLine("0");
                }
                if (checkBox4.Checked)
                {
                    sw.WriteLine("1");
                }
                else
                {
                    sw.WriteLine("0");
                }
                sw.Close();
                fs.Close();
            }
            catch
            {}
            if (button3.Enabled == false || button2.Enabled == false)
            {
                force_stop();
            }
        }

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
                {
                    checkBox2.Checked = false;
                }
                line = sr.ReadLine();
                if (line == "0")
                {
                    checkBox5.Checked = false;
                }
                line = sr.ReadLine();
                if (line == "0")
                {
                    checkBox3.Checked = false;
                }
                line = sr.ReadLine();
                if (line == "0")
                {
                    checkBox4.Checked = false;
                }
                if (sr != null) sr.Close();
            }
            else
            {
                button9_Click(sender, e);
            }
            config_check_and_recover();
            folder_check_and_recover();
        }

        private void force_stop()
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

        private void button_select_path2(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int processorCount = Environment.ProcessorCount - 1;
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

        private void button_reset_system(object sender, EventArgs e)
        {
            clear_env_data_file();
        }

        private void button_force_stop(object sender, EventArgs e)
        {
            force_stop();
        }

        private void button_select_path1(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = dialog.SelectedPath;
                to_check_paper = textBox5.Text + "\\";
            }
        }

        private void button_select_path3(object sender, EventArgs e)
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
    }
   
}
