using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Coun_PDF_Pages
{
    class Excel
    {
        string path = "";
        public _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        public int rowCount;
        public int colCount;
        public Range xlRange;

        public Excel(string path, int sheet) 
        {
            this.path = path;
            wb = excel.Workbooks.Open(path, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false); 
            ws = (Worksheet)wb.Worksheets[sheet];
            xlRange = ws.UsedRange;
            rowCount = xlRange.Rows.Count + 1;
            colCount = xlRange.Columns.Count + 1;
        }

        public void ReadCell()
        {

        }
    }
}
