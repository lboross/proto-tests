using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ChangeTest
{
    public partial class WinForm : Form
    {

        Microsoft.Office.Interop.Excel.Application _excel = null;

        public WinForm(Microsoft.Office.Interop.Excel.Application Excel )
        {
            _excel = Excel;

            InitializeComponent();
        }

        public WinForm()
        {
            InitializeComponent();
        }

        private void btn_test_Click(object sender, EventArgs e)
        {

            //Excel.Workbook WB = _excel.ActiveWorkbook as Excel.Workbook;

           /* Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            WS.Range["B4"].Value = "WF WS.Range";*/

            Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            WS.Range["C4"].Value = "WF WS.Range";

            Thread.Sleep(100);

            WS.Range["C5"].Value = "WF WS.Range 2";

            Thread.Sleep(100);

            WS.Range["C6"].Value = "WF WS.Range 3";

            Thread.Sleep(100);

            WS.Range["C7"].Value = "WF WS.Range 4";

            Thread.Sleep(100);

            WS.Range["C8"].Value = "WF WS.Range 5";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Excel.Workbook WB = _excel.ActiveWorkbook as Excel.Workbook;

            Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            WS.get_Range("B4").Value = "WF WS.get_Range";
        }
    }
}
