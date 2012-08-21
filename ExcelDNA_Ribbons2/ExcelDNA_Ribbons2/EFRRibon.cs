using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;

namespace ExcelDNA_Ribbons2
{
    [ComVisible(true)]
    public class EFRRibon : ExcelRibbon
    {
        public override string GetCustomUI(string uiName)
        {
            System.Windows.Forms.MessageBox.Show("HEY");
            XmlDocument doc = new XmlDocument();
            doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelDNA_Ribbons2.Resources.Ribbon.xml"));
            return doc.InnerXml;
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("This is a test.","Information",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        public System.Drawing.Image GetRibbonControlImage(IRibbonControl control)
        {
            MessageBox.Show("GetRibbonControlImage");
            return ExcelDNA_Ribbons2.Resources.brymck_48;
        }
    }
}
