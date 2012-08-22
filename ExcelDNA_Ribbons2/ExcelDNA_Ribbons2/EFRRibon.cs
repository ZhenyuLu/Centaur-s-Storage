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
            System.Windows.Forms.MessageBox.Show("EFRRibon Loading...", "Information",MessageBoxButtons.OK, MessageBoxIcon.Information);
            XmlDocument doc = new XmlDocument();
            doc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelDNA_Ribbons2.Resources.Ribbon.xml"));
            return doc.InnerXml;
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            Excel.Application app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            Excel.Range selection = (Excel.Range)app.Selection;
            if (selection == null)
            {
                MessageBox.Show("No data has been selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Total "+selection.Cells.Count.ToString()+" cells selected " + 
                                                " with "+selection.Rows.Count.ToString() +" rows " + 
                                                "and "+selection.Columns.Count.ToString() + " columns.","Information",
                                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public System.Drawing.Image GetRibbonControlImage(IRibbonControl control)
        {
            return ExcelDNA_Ribbons2.Resources.brymck_48;
        }
    }
}
