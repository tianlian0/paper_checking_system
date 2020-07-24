using System;
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
            //本项目所用Net Framework版本为4.6，其它版本的兼容性请自行测试
            System.Threading.Mutex instance = new System.Threading.Mutex(true, "tianlian0//paper_checking_system", out bool createdNew);
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
