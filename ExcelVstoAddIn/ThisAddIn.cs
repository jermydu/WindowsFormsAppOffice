using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelVstoAddIn
{
    public partial class ThisAddIn
    {
        //excel 应用程序对象
        Excel.Application excelApp;

        //加载项启动时 触发此事件
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //获取加载项所在的应用程序
            excelApp = Globals.ThisAddIn.Application;

            excelApp.Visible = false;
            Excel.Workbook wbk;
            //打开已经存在的工作簿文件
            wbk = excelApp.Workbooks.Open(@"D:\sourcecode\WindowsFormsAppOffice\ExcelVstoAddIn\bin\Debug\test.xlsx", System.Reflection.Missing.Value, true);
            wbk.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, @"D:\sourcecode\WindowsFormsAppOffice\ExcelVstoAddIn\bin\Debug\test.pdf") ;
        }

        //加载项退出时 触发此事件
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
