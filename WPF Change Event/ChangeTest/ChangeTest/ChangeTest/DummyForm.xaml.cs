using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace ChangeTest
{
    /// <summary>
    /// Interaction logic for DummyForm.xaml
    /// </summary>
    public partial class DummyForm : Window
    {

        Microsoft.Office.Interop.Excel.Application _excel = null;

        public DummyForm(Microsoft.Office.Interop.Excel.Application Excel)
        {
            _excel = Excel;

            InitializeComponent();
        }

        public DummyForm()
        {
            InitializeComponent();
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            button1.Focus();
            button2.Focus();
        }

        private void Window_GotFocus(object sender, RoutedEventArgs e)
        {

        }
    }
}
