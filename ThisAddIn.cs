using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        public  string palabra  { get; set; }
        public Excel.Workbook activeWorkbook;
        private  Excel.Worksheet activeWorksheet;
        private Excel.Range totalRange;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            GetFirstWord();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
       public string GetFirstWord()
        {
            // book= (Workbook)Application.ActiveWorkbook;
            //book.Close(true);
            try
            {
                this.Application.ActiveWorkbook.Close(false, missing, missing);
            }
            catch (Exception)
            {

                throw;
            }
            activeWorkbook= this.Application.Workbooks.Open(@"C:\Users\GMateusDP\Documents\SandBox\WindingNumbers.xlsx");
            activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            totalRange = activeWorksheet.UsedRange;
            Excel.Range range = activeWorksheet.get_Range("G9");
            try
            {
                palabra = range.Value;
                range.Select();
                return palabra;
            }
            catch (Exception)
            {
                return "not text";
                throw;
            }
          
        }
        public string GetReferenceWord()
        {
          activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            totalRange = activeWorksheet.UsedRange;
            Excel.Range range = activeWorksheet.get_Range("G9");
            try
            {
                palabra = range.Value;
                range.Select();
                return palabra;
            }
            catch (Exception)
            {
                return "not text";
                throw;
            }

        }

        public void ReapeatColum()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = "This text was added by using code";
        }
        public  string GetNextNumber(int i)
        {
            int index = 0;
            double value = 0;
            Excel.Range r= (Excel.Range)Application.ActiveCell;
            if (i==8)
            {
                i = r.Row - 1;
            }
            else if (i==2)
            {
                i = r.Row + 1;
            }
            Math.DivRem(i, totalRange.Rows.Count+1, out index);
            try
            {
                Excel.Range cr = activeWorksheet.get_Range(string.Concat("G", index.ToString()));
                value = cr.Value;
                cr.Select();
                palabra = value.ToString("N0");
                return palabra;
            }
            catch (Exception)
            {
                                
                throw;
                return "not text";
            }
            
        }
        public string MoveNextCell()
        {
            double value;
            Excel.Range r = (Excel.Range)Application.ActiveCell;
            int row = r.Row;
            int col = r.Column;
            r= r.get_Resize(row + 1, col);
            r.Select();
            try
            {
                value = r.Value;
                palabra = value.ToString("N0");
                return palabra;
            }
            catch (Exception)
            {
                return "not text";
                //throw;
            }
           

        }
        public string GetCurrentCell()
        {
            double value;
            Excel.Range r = (Excel.Range)Application.ActiveCell;
         //  int row = r.Row;
         //  int col = r.Column;
         //   r = r.get_Resize(row + 1, col);
         //   r.Select();
            try
            {
                value = r.Value;
                palabra = value.ToString("N0");
                r.Select();
                return palabra;
            }
            catch (Exception)
            {
                return "not text";
                //throw;
            }


        }
    }
}
