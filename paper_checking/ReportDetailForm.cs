using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using RTF_Operation;

namespace paper_checking
{
    public partial class ReportDetailForm : Form
    {
        private readonly string paperName = "";
        private readonly string dp1 = "";
        private readonly string dp2 = "";
        private readonly string dp3 = "";

        public ReportDetailForm(string p,string p1,string p2, string p3)
        {
            InitializeComponent();
            paperName = p;
            dp1 = p1;
            dp2 = p2;
            dp3 = p3;
            FormLoad();
        }

        private void FormLoad()
        {
            string paperPath = Application.StartupPath + Path.DirectorySeparatorChar + dp1 + paperName + ".txt";
            string rptPath = Application.StartupPath + Path.DirectorySeparatorChar + dp3 + paperName + ".txt";
            string rptDetailPath = Application.StartupPath + Path.DirectorySeparatorChar + dp2 + paperName + ".txt";
            if ((!File.Exists(paperPath)) || (!File.Exists(rptPath)) || (!File.Exists(rptDetailPath)))
            {
                Close();
                throw new Exception();
            }
            StringBuilder sourceFile2 = new StringBuilder("");
            //sourceFile = File.ReadAllText(paperPath, Encoding.Default)+"\r\n";
            StreamReader file2 = new StreamReader(rptDetailPath, Encoding.Default);
            string line2;
            int rnnum = 0;
            while ((line2 = file2.ReadLine()) != null)
            {
                if (line2 != "")
                {
                    sourceFile2.Append(line2 + "\r\n");
                    rnnum++;
                }
            }
            file2.Close();
            //sourceFile2 += "#####\r\n";
            sourceFile2.Append("\r\n论文原文（标红为重复部分）：\r\n");
            int startWords = sourceFile2.Length - rnnum - 2;
            StreamReader file3 = new StreamReader(paperPath, Encoding.Default);
            while ((line2 = file3.ReadLine()) != null)
            {
                if (line2.Trim() != "")
                    sourceFile2.Append(line2.Replace("@@","\r\n") + "\r\n\r\n");
            }
            file3.Close();

            StreamReader file = new StreamReader(rptPath,Encoding.Default);
            string line;

            string titletemp = "查重报告\r\n\r\n\r\n    头部           中前部          中部          中后部          尾部\r\n\r\n被测论文：" + paperName + "\r\n总重复率： " + file.ReadLine() + " %\r\n\r\n其中包含：\r\n";
            //sourceFile2 = titletemp + sourceFile2+"\r\n";
            startWords += titletemp.Length - 9;
            //startWords = sourceFile2.IndexOf("#####\r\n") + "#####\r\n".Length;
            int i = 0;
            richTextBox1.Text = titletemp + sourceFile2.ToString() + "声明：开源版查重报告较为简易，查重报告中所有内容仅供参考。商业合作、技术咨询及BUG反馈可联系微信：di18810760681，或QQ：1253332375\r\n商用版为您提供更丰富功能、更全面详尽的查重报告和更高的软件稳定性，商用试用版下载链接：https://xincheck.com/?id=28";
            //richTextBox1.Text = sourceFile2 + sourceFile;
            //int i = sourceFile2.Replace("\r\n"," ").Length;

            Image myimage = new Bitmap(575, 35);
            Graphics graphics = Graphics.FromImage(myimage);
            Point p01 = new Point(0, 1);
            Point p02 = new Point(0, 28);

            Point p1 = new Point(0, 29);
            Point p2 = new Point(0, 0);
            Point p3 = new Point(574, 0);
            Point p4 = new Point(574, 29);
            graphics.Clear(Color.White);
            graphics.DrawLine(Pens.Black, p1, p2);
            graphics.DrawLine(Pens.Black, p1, p4);
            graphics.DrawLine(Pens.Black, p3, p2);
            graphics.DrawLine(Pens.Black, p3, p4);

            p1.Y = 29;
            p2.Y = 34;
            for (int sss=0;sss<5;sss++)
            {
                p1.X = 125 * (sss + 1);
                p2.X = 125 * (sss + 1);
                graphics.DrawLine(Pens.Black, p1, p2);
            }

            int[] bucket = new int[575];

            int totalWords = int.Parse(file.ReadLine());
            while ((line = file.ReadLine()) != null)
            {
                if (line == "1")
                {
                    int index = (int)(575.0 * i / totalWords);
                    if (index > 0 && index < 574)
                    {
                        bucket[index]++;
                    }
                    richTextBox1.Select(i + startWords, 1);
                    richTextBox1.SelectionColor = Color.Red;
                }
                i++;
            }

            int minvalue = Math.Max((int)(0.3 * totalWords / 575), 1);
            for (int p = 0; p < 575; p++)
            {
                if (bucket[p] >= minvalue)
                {
                    p01.X = p;
                    p02.X = p;
                    graphics.DrawLine(Pens.Red, p01, p02);
                }
            }
            file.Close();

            Text = paperName + " 标注查重";

            richTextBox1.Select(0, "查重报告\r\n".Length);
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;

            richTextBox1.Select("查重报告\r\n".Length, 0);
            RTB_InsertImg.InsertImage(richTextBox1, myimage);

            richTextBox1.Select(1, 0);
        }
    }
}
