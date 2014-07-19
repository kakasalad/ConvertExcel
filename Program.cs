using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ConvertExcel
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.UnhandledException +=new UnhandledExceptionEventHandler(UnhandledExceptionEventHandler);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        static void UnhandledExceptionEventHandler(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                LogHelper.Info(e.ExceptionObject.ToString());//LogHelper是写日志的类，这里，可以直接写到文件里
            }
            catch
            {
            }
        }
    }
}
