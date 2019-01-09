using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xaml;

using Excel = Microsoft.Office.Interop.Excel;

namespace ChangeTest
{
    /// <summary>
    /// Interaction logic for WpfForm.xaml
    /// </summary>
    public partial class WpfForm : Window
    {

        Microsoft.Office.Interop.Excel.Application _excel = null;

        private int FocusCount = 0;

        private int ActivateCount = 0;

        public WpfForm(Microsoft.Office.Interop.Excel.Application Excel)
        {
            _excel = Excel;

            InitializeComponent();
        }

        public WpfForm()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {

            //Excel.Workbook WB = _excel.ActiveWorkbook as Excel.Workbook;

            Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            WS.Range["B4"].Value = "WPF WS.Range";

            Thread.Sleep(100);

            WS.Range["B5"].Value = "WPF WS.Range 2";

            Thread.Sleep(100);

            WS.Range["B6"].Value = "WPF WS.Range 3";

            Thread.Sleep(100);

            WS.Range["B7"].Value = "WPF WS.Range 4";

            Thread.Sleep(100);

            WS.Range["B8"].Value = "WPF WS.Range 5";

            //WS.get_Range("B4").Value = "WPF AA";

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {

            //Excel.Workbook WB = _excel.ActiveWorkbook as Excel.Workbook;

            Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;


            WS.Range["B10"].Value = "WPF WS.Range B";

            //WS.get_Range("B4").Value = "WPF WS.get_Range";

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            //Excel.Workbook WB = _excel.ActiveWorkbook as Excel.Workbook;

            //Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            //WS.Range["B4"].Value = "Range";
           // WS.get_Range("B4").Value = "get_Range";

        }

        private async Task formWait(int ms)
        {
            Thread.Sleep(ms);

            return;
        }

        private async void button3_Click(object sender, RoutedEventArgs e)
        {
            _excel.StatusBar = "Status A";

            await formWait(2000);

            _excel.StatusBar = "Status B";
        }

        private void Window_GotFocus(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("Got Focus");

            //Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            //WS.Range["E5"].Value = "FOCUS";

            lbl_status.Content = "FOCUS "+ FocusCount;
            FocusCount++;

        }

        private void Window_Activated(object sender, EventArgs e)
        {
            //MessageBox.Show("Activated");

            //Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            //WS.Range["E4"].Value = "ACTIVATED";

            lbl_status.Content = "ACTIVATED " + ActivateCount;

            ActivateCount++;
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
           /* Excel.Worksheet WS = _excel.ActiveSheet as Excel.Worksheet;

            WS.Range["C4"].Value = "WPF WS.Range";

            Thread.Sleep(100);

            WS.Range["C5"].Value = "WPF WS.Range 2";

            Thread.Sleep(100);

            WS.Range["C6"].Value = "WPF WS.Range 3";

            Thread.Sleep(100);

            WS.Range["C7"].Value = "WPF WS.Range 4";

            Thread.Sleep(100);

            WS.Range["C8"].Value = "WPF WS.Range 5";*/
        }
    }
}
