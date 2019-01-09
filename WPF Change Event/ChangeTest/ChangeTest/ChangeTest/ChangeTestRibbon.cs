using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading;
using System.Windows.Threading;

using Microsoft.Office.Interop.Excel;

namespace ChangeTest
{
    public partial class ChangeTestRibbon
    {
        private int EventCount = 1;

        public void workbook_Change(Object sh, Range Target)
        {
            //MessageBox.Show("Book Change HIT", "Book Change: " + (sh as Worksheet).Name);

            try
            {
                Worksheet WS = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                Globals.ThisAddIn.Application.EnableEvents = false;

                WS.Range["A1"].Value = "Event: " + EventCount;

                Globals.ThisAddIn.Application.EnableEvents = true;

                EventCount++;
            }
           catch( Exception ex)
            {

            }

            return;
        }

        public static void worksheet_Change( Range Target)
        {
            MessageBox.Show("Sheet Change HIT", "Sheet Change: ");
            return;
        }

        private void ChangeTestRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// Adds the workbook change event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            wb.SheetChange += workbook_Change;

            MessageBox.Show("Workbook SheetChange Registered", "OK");
        }

        /// <summary>
        /// Adds the worksheet change event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            ws.Change += worksheet_Change;

            MessageBox.Show("Worksheet Change Registered", "OK");
        }

        /// <summary>
        /// Opens the Windows Form from the Main Thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            WinForm wf = new WinForm(Globals.ThisAddIn.Application);

            wf.Show();
        }

        /// <summary>
        /// Opens the Windows Form in a background thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            object[] options = new object[] { Globals.ThisAddIn.Application };
            WindowHelper.CreateFormWithOptions<WinForm>(700, 550, 345, 250, options, (IntPtr)Globals.ThisAddIn.Application.Hwnd);
        }

        /// <summary>
        /// Opens the WPF Form from the Main Thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            WpfForm ws = new WpfForm(Globals.ThisAddIn.Application);

            ws.Show();
        }

        /// <summary>
        /// Opens the WPF form in a background thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            object[] options = new object[] { Globals.ThisAddIn.Application };
            WindowHelper.CreateWindowWithOptions<WpfForm>(700, 550, 345, 250, options, (IntPtr)Globals.ThisAddIn.Application.Hwnd);

        }

        /// <summary>
        /// An attempt to work-around the issue by setting the event handler, breaking it in the background thread, and re-applying the handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            //Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            //ws.Change += worksheet_Change;

            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            wb.SheetChange += workbook_Change;

            Thread newWindowThread = new Thread(new ThreadStart(() =>
            {
                SynchronizationContext.SetSynchronizationContext(
                    new DispatcherSynchronizationContext(
                        Dispatcher.CurrentDispatcher));
                try
                {

                    //Globals.ThisAddIn.Application.EnableEvents = false;

                    DummyForm df = new DummyForm(Globals.ThisAddIn.Application);

                    //df.Activate();

                    //df.Hide();

                    df.ShowDialog();

                    //df.Activate();

                    //df.Close();

                    //df.ShowDialog();

                    //Globals.ThisAddIn.Application.EnableEvents = true;

                    wb.SheetChange += workbook_Change;

                }
                catch (Exception ex)
                {
                    
                }
            }));

            newWindowThread.Name = "Test";
            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
            //newWindowThread.Abort();



        }

        /// <summary>
        /// Opens the WPF Form from the Main Thread as a Dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            WpfForm ws = new WpfForm(Globals.ThisAddIn.Application);

            ws.ShowDialog();
        }

        /// <summary>
        /// Updates a cell value from the Ribbon (Main thread)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet WS = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            WS.Range["B4"].Value = "RBN AA";

            //WS.get_Range("B4").Value = "RBN BB";


        }
    }
}
