using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;


namespace PowerPointVstoAddIn
{
    public partial class ThisAddIn
    {
        PowerPoint.Application appPPT;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //获取加载项所在的应用程序
            appPPT = Globals.ThisAddIn.Application;

            PowerPoint.Presentation document = appPPT.Presentations.Open(@"D:\sourcecode\WindowsFormsAppOffice\PowerPointVstoAddIn\bin\Debug\test.pptx", ReadOnly: Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse,WithWindow: Office.MsoTriState.msoFalse);
            
            document.ExportAsFixedFormat(@"D:\sourcecode\WindowsFormsAppOffice\PowerPointVstoAddIn\bin\Debug\test.pdf", PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
        }

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
