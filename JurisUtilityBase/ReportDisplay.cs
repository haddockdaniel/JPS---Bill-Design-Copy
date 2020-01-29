using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JurisUtilityBase
{
    public partial class ReportDisplay : Form
    {
        public ReportDisplay(DataSet time, DataSet expense)
        {
            InitializeComponent();
            dataGridView1.DataSource = time.Tables[0];
            dataGridView2.DataSource = expense.Tables[0];
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {

            //PrinterDialog pd = new PrinterDialog();
          //  pd.ShowDialog();
          //  string printer = pd.printerName;
          //  if (!string.IsNullOrEmpty(printer))
           // {


                saveClient();
                saveMatter();
           // }
        }

        private void saveClient()
        {
            Cursor.Current = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "Client";

            int StartCol = 1;
            int StartRow = 1;
            int j = 0, i = 0;

            //Write Headers
            for (j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }

            StartRow++;

            //Write datagridview content
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        myRange.Value2 = myRange.Value2.Trim();
                    }
                    catch
                    {
                        ;
                    }
                }
            }

            Microsoft.Office.Interop.Excel.Range usedrange = xlWorkSheet.UsedRange;
            usedrange.Columns.AutoFit();
            xlApp.Visible = false;
            var _with1 = xlWorkSheet.PageSetup;
            _with1.Zoom = false;
            _with1.PrintGridlines = true;
            _with1.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
            _with1.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            _with1.FitToPagesWide = 1;
            _with1.FitToPagesTall = false;

            _with1.PrintTitleRows = "$1:$" + dataGridView1.Columns.Count.ToString();


            SaveFileDialog ssv = new SaveFileDialog();
            ssv.Filter = "Excel File|*.xlsx";
            ssv.DefaultExt = "xlsx"; 
            ssv.Title = "Save Client Error File";
            ssv.ShowDialog();
            string path = ssv.FileName;

            xlWorkBook.SaveAs(path, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //string Defprinter = null;
            //Defprinter = xlApp.ActivePrinter;
            //xlApp.ActivePrinter = printer;

            // Print the range
           // usedrange.PrintOutEx(misValue, misValue, misValue, misValue,
            //misValue, misValue, misValue, misValue);
            // }
            //xlApp.ActivePrinter = Defprinter;

            // Cleanup:
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(xlWorkSheet);

            xlWorkBook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkBook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            Cursor.Current = Cursors.Default;
        }

        private void saveMatter()
        {
            Cursor.Current = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "Matter";

            int StartCol = 1;
            int StartRow = 1;
            int j = 0, i = 0;

            //Write Headers
            for (j = 0; j < dataGridView2.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView2.Columns[j].HeaderText;
            }

            StartRow++;

            //Write datagridview content
            for (i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView2[j, i].Value == null ? "" : dataGridView2[j, i].Value;
                        myRange.Value2 = myRange.Value2.Trim();
                    }
                    catch
                    {
                        ;
                    }
                }
            }

            Microsoft.Office.Interop.Excel.Range usedrange = xlWorkSheet.UsedRange;
            usedrange.Columns.AutoFit();
            xlApp.Visible = false;
            var _with1 = xlWorkSheet.PageSetup;
            _with1.Zoom = false;
            _with1.PrintGridlines = true;
            _with1.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
            _with1.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            _with1.FitToPagesWide = 1;
            _with1.FitToPagesTall = false;

            _with1.PrintTitleRows = "$1:$" + dataGridView2.Columns.Count.ToString();

            SaveFileDialog ssv = new SaveFileDialog();
            ssv.Filter = "|Excel File|*.xlsx";
            ssv.DefaultExt = "xlsx";
            ssv.Title = "Save Matter Error File";
            ssv.ShowDialog();
            string path = ssv.FileName;

            xlWorkBook.SaveAs(path, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //string Defprinter = null;
           // Defprinter = xlApp.ActivePrinter;
           // xlApp.ActivePrinter = printer;

            // Print the range
            //usedrange.PrintOutEx(misValue, misValue, misValue, misValue,
           // misValue, misValue, misValue, misValue);
            // }
            //xlApp.ActivePrinter = Defprinter;

            // Cleanup:
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(xlWorkSheet);

            xlWorkBook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(xlWorkBook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            Cursor.Current = Cursors.Default;
        }




    }
}
