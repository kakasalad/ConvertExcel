using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Runtime.InteropServices;

namespace ConvertExcel
{
    public partial class Form1 : Form
    {
        public static string msg = "";
        private delegate void FlushClient(); //代理
        Microsoft.Office.Interop.Excel.ApplicationClass myapp;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (this.txtOrg.Text.Trim() == string.Empty || this.txtNew.Text.Trim() == string.Empty)
            {
                MessageBox.Show("把路径填上啊魂淡！！！", "魂淡", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (!Directory.Exists(this.txtOrg.Text.Trim()) || !Directory.Exists(this.txtNew.Text.Trim()))
            {
                MessageBox.Show("路径不对啊魂淡！！！", "魂淡", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                Thread t = new Thread(Do);
                t.IsBackground = true;
                t.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "出错啦！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Do()
        {
            DirectoryInfo folder = new DirectoryInfo(this.txtOrg.Text);
            GetFiles(folder);
            msg = string.Format("全部转换结束！");
            AppendMsg();
        }

        private void GetFiles(DirectoryInfo folder)
        {
            try
            {
                foreach (FileInfo file in folder.GetFiles())
                {
                    if (file.Extension != ".xls") continue;
                    myapp = new ApplicationClass();
                    Workbooks workbooks = myapp.Workbooks;
                    Workbook workbook = workbooks.Open(file.FullName,
                        Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    workbook.SaveAs(Path.Combine(this.txtNew.Text, file.Name.Substring(0, file.Name.LastIndexOf('.')) + ".xlsx"), Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    msg = string.Format(file.FullName + "----转换完毕！");
                    AppendMsg();
                    KillProcess(myapp.Hwnd);
                    GC.Collect();
                }
                foreach (DirectoryInfo di in folder.GetDirectories())
                {
                    GetFiles(di);
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                AppendMsg();
            }
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        /// <summary>
        /// 杀掉进程
        /// </summary>
        /// <param name="ID"></param>
        private static void KillProcess(int ID)
        {
            try
            {
                IntPtr t = new IntPtr(ID);
                int k = 0;
                GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
                if (k != 0)
                {
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                    p.Kill();     //关闭进程k
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "出错啦！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string Convert(string path, string filename)
        {
            OleDbConnection conn;
            string strConn = string.Empty;
            try
            {
                try
                {
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=No;IMEX=1'";
                    conn = new OleDbConnection(strConn);
                    conn.Open();
                }
                catch
                {
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
                    conn = new OleDbConnection(strConn);
                    conn.Open();
                }
            }
            catch (Exception ex)
            {
                return "打开文件失败!请在服务器安装Microsoft Access Database Engine 2010 Redistributable";
            }

            List<string> TableNames = new List<string>();
            System.Data.DataTable dttable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow item in dttable.Rows)
                TableNames.Add(item["TABLE_NAME"] as string);

            if (TableNames.Count == 0) return "excel里没数据";

            OleDbDataAdapter da = new OleDbDataAdapter(string.Format("select * from [{0}]", TableNames[0]), conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            //填充新的xlsx文件
            try
            {
                SuperToExcel(dt, Path.Combine(this.txtNew.Text, filename + ".xlsx"));
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                AppendMsg();
            }
            return string.Empty;
        }

        //高效导出Excel
        public static bool SuperToExcel(DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                app.Visible = false;

                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                Microsoft.Office.Interop.Excel.Range range;
                int colIndex = 0;
                int rowIndex = 0;
                int colCount = excelTable.Columns.Count;
                int rowCount = excelTable.Rows.Count;

                //创建缓存数据
                object[,] objData = new object[rowCount + 1, colCount];

                //写标题
                int size = excelTable.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelTable.Columns[i].Caption;
                }
                range = (Range)wSheet.get_Range(app.Cells[1, 1], app.Cells[1, colCount]);
                range.Interior.ColorIndex = 15;//背景色 灰色
                range.Font.Bold = true;//粗字体
                //获取实际数据

                for (rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    for (colIndex = 0; colIndex < colCount; colIndex++)
                    {
                        objData[rowIndex, colIndex] = excelTable.Rows[rowIndex][colIndex].ToString();
                    }
                }


                //写入Excel 
                range = (Range)wSheet.get_Range(app.Cells[2, 1], app.Cells[rowCount + 1, colCount]);
                range.NumberFormatLocal = "@";//设置数字文本格式
                range.Value2 = objData;
                //Application.DoEvents();

                wSheet.Columns.AutoFit();

                //设置禁止弹出保存和覆盖的询问提示框 
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;

                wBook.Saved = true;
                wBook.SaveCopyAs(filePath);

                app.Quit();
                app = null;
                GC.Collect();
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
            }
        }


        private void AppendMsg()
        {
            if (this.txtResult.InvokeRequired)//等待异步
            {
                FlushClient fc = new FlushClient(AppendMsg);
                this.Invoke(fc); //通过代理调用刷新方法
            }
            else
            {
                if (!string.IsNullOrEmpty(msg))
                {
                    this.txtResult.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss：") + msg);
                    this.txtResult.AppendText("\r\n");
                    this.txtResult.ScrollToCaret();
                }
            }
        }

        private void txtOrg_DoubleClick(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.txtOrg.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void txtNew_DoubleClick(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.txtNew.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
