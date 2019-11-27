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

            System.Threading.Mutex instance = new System.Threading.Mutex(true, "CachongSingleStart0.9", out bool createdNew);
            if (createdNew)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                try
                {
                    Application.Run(new MainForm());
                }
                catch (Exception)
                { }
                finally
                {
                    if (instance != null)
                    {
                        instance.ReleaseMutex();
                        instance.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("已经存在一个正在运行的查重程序！");
                Application.Exit();
            }

        }
    }
}
