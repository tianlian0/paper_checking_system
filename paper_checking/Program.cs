using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace paper_checking
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]

        static void Main()
        {

            Boolean createdNew;
            System.Threading.Mutex instance = new System.Threading.Mutex(true, "CachongSingleStart0.8", out createdNew); //同步基元变量 
            if (createdNew)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                try
                {
                    Application.Run(new Form1());
                    instance.ReleaseMutex();
                }
                catch (Exception)
                { }
            }
            else
            {
                MessageBox.Show("已经存在一个正在运行的查重程序！");
                Application.Exit();
            }

        }
    }
}
