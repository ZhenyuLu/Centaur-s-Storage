﻿using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using System.Data;

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
                DialogResult dr = MessageBox.Show("First row is the header title?","Question",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                DataTable dt = new DataTable();
                if (dr == DialogResult.Yes)
                {
                    dt = StoreRange2DataTable(selection, true);
                }
                else
                {
                    dt = StoreRange2DataTable(selection, false);
                }
            }
        }

        /// <summary>
        /// Save the data of active cells to a data table
        /// </summary>
        /// <param name="range">range of data</param>
        /// <param name="firstRowHeader">true means the first row is the header string, false otherwise.</param>
        /// <returns>a datatable containing the data</returns>
        public DataTable StoreRange2DataTable(Excel.Range range, bool firstRowHeader)
        {
            DataTable dt = new DataTable();
            object[,] temp = (object[,])range.Cells.Value2;
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;

            try
            {
                if (firstRowHeader)
                {
                    for (int i = 1; i <= colCount; i++)
                    {
                        dt.Columns.Add(temp[1,i].ToString());
                    }

                    for (int i = 1; i <= rowCount; i++)
                    {
                        DataRow row = dt.NewRow();
                        for (int j = 1; j <= colCount; j++)
                        {
                            row[j-1] = temp[i, j].ToString();
                        }
                        dt.Rows.Add(row);
                    }
                }
                else
                {
                    for (int i = 1; i <= colCount; i++)
                    {
                        dt.Columns.Add("var" + (i + 1).ToString());
                    }
                    for (int i = 1; i <= rowCount; i++)
                    {
                        DataRow row = dt.NewRow();
                        for (int j = 1; j < colCount; j++)
                        {
                            row[j-1] = temp[i,j].ToString();
                        }
                        dt.Rows.Add(row);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Unable to pull the data out because " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            return dt;
        }

        public System.Drawing.Image GetRibbonControlImage(IRibbonControl control)
        {
            return ExcelDNA_Ribbons2.Resources.brymck_48;
        }
    }
}
