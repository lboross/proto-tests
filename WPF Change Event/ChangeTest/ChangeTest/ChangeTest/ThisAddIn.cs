using System;
using System.Reflection;
using System.Linq.Expressions;

using Expression = System.Linq.Expressions.Expression;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Windows.Forms;

/// <summary>
/// Test application to highlight issues with Workbook.SheetChange event 
/// </summary>
namespace ChangeTest
{

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            //MessageBox.Show("Workbook Activate", "Activate");

            //Wb.SheetChange += ChangeTestRibbon.workbook_Change;
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
    }

    public class WindowBase : Window
    {
        protected ManualResetEvent _windowCloseEvent;

        public WindowBase()
        {
            DataContextChanged += new DependencyPropertyChangedEventHandler(OnDataContextChanged);
        }

        public void InitializeCloseAction(ManualResetEvent windowCloseEvent)
        {
            _windowCloseEvent = windowCloseEvent;

            var worker = new BackgroundWorker();
            worker.DoWork += (o, ea) =>
            {
                while (true)
                {
                    if (_windowCloseEvent.WaitOne(0))
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            Close();
                        }));
                    }
                    Thread.Sleep(100);
                }
            };
            worker.RunWorkerAsync();
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
        }

        public void CloseReturnSetup(object arg)
        {
            _windowCloseEvent.Set();
        }
    }

    /// <summary>
    /// Helper class to create Windows
    /// </summary>
    public static class WindowHelper
    {
        public delegate T ObjectActivator<T>(params object[] args);

        public static void CreateWindow<TWindow>(int x, int y, int width, int height, IntPtr Hwnd,
            ManualResetEvent windowReadyEvent = null, ManualResetEvent windowCloseEvent = null) where TWindow : Window, new()
        {

            // Create a thread
            Thread newWindowThread = new Thread(new ThreadStart(() =>
            {
                SynchronizationContext.SetSynchronizationContext(
                    new DispatcherSynchronizationContext(
                        Dispatcher.CurrentDispatcher));
                TWindow tempWindow = new TWindow();
                WindowBase wb = tempWindow as WindowBase;
                if (wb != null && windowCloseEvent != null)
                {
                    wb.InitializeCloseAction(windowCloseEvent);
                }

                // Get hWnd for non-WPF window
                var ownerWindowHandle = (IntPtr)Hwnd;

                // Set the owned WPF window’s owner with the non-WPF owner window
                var helper = new WindowInteropHelper(tempWindow);
                helper.Owner = ownerWindowHandle;

                tempWindow.Closed +=
                    (s, e) =>
                    {
                        windowCloseEvent.Set();
                        Dispatcher.CurrentDispatcher.InvokeShutdown();
                    };

                if (windowReadyEvent != null)
                    windowReadyEvent.Set();
                tempWindow.Show();

 
                System.Windows.Threading.Dispatcher.Run();
            }));

            newWindowThread.SetApartmentState(ApartmentState.STA);
            newWindowThread.IsBackground = true;
            newWindowThread.Start();
        }


        public static void CreateFormWithOptions<TWindow>(int x, int y, int width, int height, object[] options, IntPtr Hwnd) where TWindow : Form, new()
        {
            try
            {
                // Create a thread
                Thread newWindowThread = new Thread(new ThreadStart(() =>
                {
                    SynchronizationContext.SetSynchronizationContext(
                        new DispatcherSynchronizationContext(
                            Dispatcher.CurrentDispatcher));

                    //Note: This assumes the parametered consructor is defined first in the class
                    ConstructorInfo ctor = (typeof(TWindow)).GetConstructors()[0];
                    ObjectActivator<TWindow> createdActivator = GetActivator<TWindow>(ctor);
                    TWindow tempWindow = createdActivator(options);
                    tempWindow.Show();

                    try
                    {
                        System.Windows.Threading.Dispatcher.Run();
                    }
                    catch (Exception ex)
                    {
                        tempWindow.Close();
                    }
                }));

                newWindowThread.Name = typeof(TWindow).Name;
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void CreateWindowWithOptions<TWindow>(int x, int y, int width, int height, object[] options, IntPtr Hwnd,
           ManualResetEvent windowReadyEvent = null, ManualResetEvent windowCloseEvent = null) where TWindow : Window, new()
        {
            try
            {
                // Create a thread
                Thread newWindowThread = new Thread(new ThreadStart(() =>
                {
                    SynchronizationContext.SetSynchronizationContext(
                        new DispatcherSynchronizationContext(
                            Dispatcher.CurrentDispatcher));

                    //Note: This assumes the parametered consructor is defined first in the class
                    ConstructorInfo ctor = (typeof(TWindow)).GetConstructors()[0];
                    ObjectActivator<TWindow> createdActivator = GetActivator<TWindow>(ctor);
                    TWindow tempWindow = createdActivator(options);
                    tempWindow.Show();

                    WindowBase wb = tempWindow as WindowBase;
                    if (wb != null && windowCloseEvent != null)
                    {
                        wb.InitializeCloseAction(windowCloseEvent);
                    }

                    // Get hWnd for non-WPF window
                    var ownerWindowHandle = (IntPtr)Hwnd;

                    // Set the owned WPF window’s owner with the non-WPF owner window
                    var helper = new WindowInteropHelper(tempWindow);
                    helper.Owner = ownerWindowHandle;

                    tempWindow.Closed +=
                        (s, e) =>
                        {
                            if (windowCloseEvent != null)
                                windowCloseEvent.Set();

                            Dispatcher.CurrentDispatcher.InvokeShutdown();
                        };

                    if (windowReadyEvent != null)
                        windowReadyEvent.Set();

                    try
                    {
                        System.Windows.Threading.Dispatcher.Run();
                    }
                    catch (Exception ex)
                    {
                        tempWindow.Close();
                    }
                }));

                newWindowThread.Name = typeof(TWindow).Name;
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static ObjectActivator<T> GetActivator<T>(ConstructorInfo ctor)
        {
            Type type = ctor.DeclaringType;
            ParameterInfo[] paramsInfo = ctor.GetParameters();

            //create a single param of type object[]
            ParameterExpression param =
                Expression.Parameter(typeof(object[]), "args");

            Expression[] argsExp =
                new Expression[paramsInfo.Length];

            //pick each arg from the params array 
            //and create a typed expression of them
            for (int i = 0; i < paramsInfo.Length; i++)
            {
                Expression index = Expression.Constant(i);
                Type paramType = paramsInfo[i].ParameterType;

                Expression paramAccessorExp =
                    Expression.ArrayIndex(param, index);

                Expression paramCastExp =
                    Expression.Convert(paramAccessorExp, paramType);

                argsExp[i] = paramCastExp;
            }

            //make a NewExpression that calls the
            //ctor with the args we just created
            NewExpression newExp = Expression.New(ctor, argsExp);

            //create a lambda with the New
            //Expression as body and our param object[] as arg
            LambdaExpression lambda =
                Expression.Lambda(typeof(ObjectActivator<T>), newExp, param);

            //compile it
            ObjectActivator<T> compiled = (ObjectActivator<T>)lambda.Compile();
            return compiled;
        }
    }

}
