using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace paper_checking
{
    public partial class ReportListForm : Form
    {
        string dp1 = "", dp2 = "", dp3 = "";
        int curPageNum = 0;
        int sumPageNum = 0;
        public ReportListForm(string p1, string p2, string p3)
        {
            InitializeComponent();
            dp1 = p1;
            dp2 = p2;
            dp3 = p3;
            Form3_Load2();
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                int Index = listView1.SelectedItems[0].Index;
                String aa = listView1.Items[Index].Text;
                ReportDetailForm f = new ReportDetailForm(aa, dp1, dp2, dp3);
                f.Show();
            }
        }

        private void ListView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = listView1.HitTest(e.X, e.Y);
            if (info.Item != null)
            {
                if (listView1.SelectedItems.Count > 0)
                {
                    int Index = listView1.SelectedItems[0].Index;
                    String aa = listView1.Items[Index].Text;
                    ReportDetailForm f = new ReportDetailForm(aa, dp1, dp2, dp3);
                    f.Show();
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int i=0;
            listView1.SelectedItems.Clear();
            for (; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].Text.Contains(textBox1.Text))
                {
                    listView1.Items[i].Selected = true;
                }
            }
            listView1.Focus();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (curPageNum <= sumPageNum)
            {
                return;
            }

            listView1.Clear();
            for (int FileInfoNo = curPageNum * 200-200; FileInfoNo < curPageNum * 200; FileInfoNo++)
            {
                FileInfo NextFile = fileInfo[FileInfoNo];

                FileStream fs = new FileStream(NextFile.FullName, FileMode.Open, FileAccess.Read);
                StreamReader read = new StreamReader(fs, Encoding.Default);
                string strReadline = read.ReadLine();
                fs.Close();
                read.Close();

                ListViewItem litem = new ListViewItem();
                litem.Text = NextFile.Name.Substring(0, NextFile.Name.Length - 8);
                litem.SubItems.Add(strReadline);
                listView1.Items.Add(litem);
            }
            curPageNum--;
            //label3.Text = (curPageNum - 1).ToString();
        }

        FileInfo[] fileInfo;
        private void Button3_Click(object sender, EventArgs e)
        {
            if (curPageNum >= sumPageNum)
            {
                return;
            }
            int k;
            if (fileInfo.Length <= 200 + curPageNum * 200)
                k = fileInfo.Length;
            else
                k = 200 + curPageNum * 200;

            listView1.Clear();
            for (int FileInfoNo = curPageNum * 200; FileInfoNo < k; FileInfoNo++)
            {
                FileInfo NextFile = fileInfo[FileInfoNo];

                FileStream fs = new FileStream(NextFile.FullName, FileMode.Open, FileAccess.Read);
                StreamReader read = new StreamReader(fs, Encoding.Default);
                string strReadline = read.ReadLine();
                fs.Close();
                read.Close();

                ListViewItem litem = new ListViewItem();
                litem.Text = NextFile.Name.Substring(0, NextFile.Name.Length - 8);
                litem.SubItems.Add(strReadline);
                listView1.Items.Add(litem);
            }
            curPageNum++;
            //label3.Text = (curPageNum + 1).ToString();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            int i = 0;
            listView1.SelectedItems.Clear();
            for (; i < listView1.Items.Count; i++)
            {
                if (float.Parse(listView1.Items[i].SubItems[1].Text)> float.Parse(textBox2.Text))
                {
                    listView1.Items[i].Selected = true;
                }
            }
            listView1.Focus();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void Form3_Load2()
        {
            DirectoryInfo sourceFolder = new DirectoryInfo(dp3);
            if (!sourceFolder.Exists)
            {
                MessageBox.Show("Data文件夹丢失，无法显示查重结果！");
                Close();
                return;
            }
            fileInfo = sourceFolder.GetFiles();
            if (fileInfo.Length % 200 == 0)
                sumPageNum = fileInfo.Length / 200;
            else
                sumPageNum = fileInfo.Length / 200 + 1;
            
            //if (sumPageNum > 0)
            //{
            //    sumPageNum--;
            //}
            //else
            //{
            //    label4.Text = (sumPageNum).ToString();
            //}


            int k;
            if (fileInfo.Length <= 3000)
                k = fileInfo.Length;
            else
                k = 3000;

            for (int FileInfoNo = 0; FileInfoNo < k; FileInfoNo++)
            {
                FileInfo NextFile = fileInfo[FileInfoNo];

                FileStream fs = new FileStream(NextFile.FullName, FileMode.Open, FileAccess.Read);
                StreamReader read = new StreamReader(fs, Encoding.Default);
                string strReadline = read.ReadLine();
                fs.Close();
                read.Close();

                ListViewItem litem = new ListViewItem();
                litem.Text = NextFile.Name.Substring(0, NextFile.Name.Length - 4);
                litem.SubItems.Add(strReadline);
                listView1.Items.Add(litem);
            }
        }
    }
}
