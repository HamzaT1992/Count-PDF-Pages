using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Coun_PDF_Pages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ppath = openFileDialog1.FileName;
            int colNB = (int)numericUpDown1.Value;
            Excel xlApp = new Excel(ppath, 1);
            object misValue = System.Reflection.Missing.Value;

            for (int i = 2; i <= xlApp.rowCount; i++ )
            {
                Range cell = xlApp.xlRange.Cells[i, colNB] as Range;
                if (cell.Value2 == null)
                    continue;
                string pdfPath = cell.Value2.ToString();
                PdfReader pdfReader = new PdfReader(pdfPath);
                xlApp.ws.Cells[i, colNB + 1] = pdfReader.NumberOfPages;
            }
            // Disable file override confirmaton message  
            xlApp.excel.DisplayAlerts = false;
            xlApp.wb.SaveAs(ppath, _Excel.XlFileFormat.xlOpenXMLWorkbook,
                misValue, misValue, misValue, misValue, _Excel.XlSaveAsAccessMode.xlNoChange,
                _Excel.XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue,
                misValue, misValue);
            xlApp.wb.Close();
            xlApp.excel.Quit();
            System.Diagnostics.Process.Start(ppath);

            
            label4.Text = "\nRecords Added successfully...";
        }
    }
}
