using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//引入office excel
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WindowsFormsAppOffice
{
    public partial class Form1 : Form
    {
        //申明excel应用程序对象
        Excel.Application ExcelApp;
        String result;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //操作 application 对象
            CreateExcelApp();
            GetExcelApp();
            AppCommonEvent();
            AppCommonFun();
        }

        #region Application 相关操作
        //获取excel对象
        public void GetExcelApp()
        {
            //获取正在运行的excel
            ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            MessageBox.Show(ExcelApp.Version + "");
        }

        //创建excel对象
        public void CreateExcelApp()
        {
            Excel.Application newApp = new Excel.Application();
            newApp.Visible = true;
            newApp.Caption = "new app";
        }

        //application 常用方法
        public void AppCommonFun()
        {
            ExcelApp.Undo();                //撤销 相当于CTRL+Z
            ExcelApp.Workbooks.Close();     //关闭所有工作簿
            ExcelApp.Quit();                //退出Excel
        }

        //application 常用事件
        public void AppCommonEvent()
        {
            ExcelApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(ExcelAppWorkbookBeforeClose);
            ExcelApp.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(ExcelAppWorkbookBeforeSelectionChange);
        }

        //工作簿关闭之前触发此事件
        public void ExcelAppWorkbookBeforeClose(Excel.Workbook wbk, ref bool Cancel)
        {
            MessageBox.Show("即将关闭" + wbk.FullName);
            Cancel = true;      //取消关闭
        }

        //工作簿单元格所选区域发生变化时触发此事件
        public void ExcelAppWorkbookBeforeSelectionChange(object sh, Excel.Range tagRange)
        {
            //鼠标选中区域颜色变为黄色
            tagRange.Interior.Color = System.Drawing.Color.Yellow;
        }

        //Application 集合
        public void AppCollection()
        {
            //excel内所有内置对话框个数
            MessageBox.Show(ExcelApp.Dialogs.Count + "");

            //遍历所有加载项信息
            foreach (Excel.AddIn adn in ExcelApp.AddIns)
            {
                result += (adn.Name + "\t" + adn.Installed + "\n");
            }

            //工具栏个数
            MessageBox.Show(ExcelApp.CommandBars.Count + "");
            //遍历工具栏信息
            foreach (Office.CommandBar cmd in ExcelApp.CommandBars)
            {
                result += (cmd.Name + "\t" + cmd.Type + "\n");
            }
        }
        #endregion

        #region workbook 相关操作

        //workbook 常用属性
        public void WorkbookCommonProperty()
        {
            Excel.Workbook wbk;
            //活动工作簿
            wbk = ExcelApp.ActiveWorkbook;
            //第一个工作簿
            wbk = ExcelApp.Workbooks[1];
            //使用名称确定工作簿
            wbk = ExcelApp.Workbooks["1111.xlsx"];

            //遍历工作簿
            foreach(Excel.Workbook wk in ExcelApp.Workbooks)
            {
                result += (wk.Name + "\n");
            }

            //工作簿是否已经保存
            bool sv = wbk.Saved;
            //工作簿文件路径
            string path = wbk.Path;
            //是否有密码保护
            bool pwd = wbk.HasPassword;
        }

        //workbook 常用方法
        public void WorkbookCommonFun()
        {
            Excel.Workbook wbk;
            //打开已经存在的工作簿文件
            wbk = ExcelApp.Workbooks.Open(@"D:\test.xlsx");

            //新建工作簿
            wbk = ExcelApp.Workbooks.Add();

            //保存工作簿
            wbk.SaveAs(@"D:\test.xlsx");

            //关闭工作簿
            wbk.Close(false,Type.Missing, Type.Missing);
        }

        //workbook 常用事件
        public void WorkbookCommonEvent()
        {
            Excel.Workbook wbk;
            wbk = ExcelApp.ActiveWorkbook;
            //当保存工作簿时触发此事件
            wbk.BeforeSave += new Excel.WorkbookEvents_BeforeSaveEventHandler(wbk_BeforeSave);
        }

        public void wbk_BeforeSave(bool UI ,ref bool Cancel)
        {
            MessageBox.Show("取消保存！");
            Cancel = true;
        }
        
        //workbook 集合
        public void WorkbookCollection()
        {
            Excel.Workbook wbk = ExcelApp.ActiveWorkbook;

            //遍历所有工作表
            foreach(Excel.Worksheet sheet in wbk.Worksheets)
            {
                result += sheet.Name + "\n";
            }

            //改变工作簿各个窗口的标题文字
            foreach(Excel.Window w in wbk.Windows)
            {
                w.Caption = "change Caption";
            }
        }
        #endregion

        //ppt 操作
        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Application appPPT;
            appPPT = new PowerPoint.Application();

            PowerPoint.Presentation document = appPPT.Presentations.Open(@"D:\sourcecode\WindowsFormsAppOffice\PowerPointVstoAddIn\bin\Debug\test.pptx", ReadOnly: Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, WithWindow: Office.MsoTriState.msoFalse);

            document.ExportAsFixedFormat(@"D:\sourcecode\WindowsFormsAppOffice\PowerPointVstoAddIn\bin\Debug\test.pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            appPPT.Quit();
        }
    }
}