using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Reflection;
using System.Diagnostics;
using Ini;

namespace Wealthy_RPT
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    /// 
    internal delegate void Invoker();

    public partial class App : Application
    {
        public App()
        {
            // Populate global variables
            LoadAppVariables gen = new LoadAppVariables();
            gen.LoadApplication();

            ApplicationInitialize = _applicationInitialize;
        }
        public static new App Current
        {
            get { return Application.Current as App; }
        }

        internal delegate void ApplicationInitializeDelegate(SplashScreen splashWindow);

        internal ApplicationInitializeDelegate ApplicationInitialize;

        private void _applicationInitialize(SplashScreen splashWindow)
        {

            // Step through usual opening processes
            splashWindow.SetProgress("Local Variables Loaded");
            Thread.Sleep(1500);

            splashWindow.SetProgress("Checking Version");
            // Check version numbers
            Global.CheckVersions();
            Thread.Sleep(1500);

            splashWindow.SetProgress("Checking Connection");
            // Check if db is online
            #region DB Online
            if (!Global.IsServerConnected(Global.ConnectionString))
            {
                MessageBox.Show(Global.ApplicationName + " is not available.  Please try again later.  If the problem persists, please report the error to the IT Helpdesk", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Error);
                System.Threading.Thread.CurrentThread.Abort();
                Application.Current.Shutdown();
                Environment.Exit(0);
            }
            #endregion DB Online
            Thread.Sleep(1500);

            splashWindow.SetProgress("Checking User");
            // Get user details            
            User user = new User();
            Global.FullName = user.FullName;
            Global.PID = user.PID;
            Global.AccessLevel = user.AccessLevel;
            Global.Admin = user.Admin;
            Global.BA_Admin = user.BA_Admin;
            
            Global.Pop_Code_Name = user.Pop_Code_Name;
            Thread.Sleep(1500);

            splashWindow.SetProgress("Checking Annual Update");
            Thread.Sleep(1000);
            if (Global.AnnualUpdate != DateTime.Now.Year.ToString())
            {
                splashWindow.SetProgress("Performing Annual Update");
                if (PerformAnnualUpdate() == false)
                {
                    MessageBox.Show("Failure to perform annual update.  Please report to R&I.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Error);
                }
                Thread.Sleep(1000);
            }

            splashWindow.SetProgress("Checking Daily Recalculation");
            Thread.Sleep(1000);
            if (Global.DailyRecalc != DateTime.Now.ToString("dd/MM/yyyy"))
            {
                //MessageBox.Show( Global.DailyRecalc + ":" + DateTime.Now.ToString("dd/MM/yyyy"), Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
                splashWindow.SetProgress("Performing Daily Recalculation");
                if (PerformDailyRecalc()==false)
                {
                    MessageBox.Show("Failure to perform daily recalculation.  Please report to R&I.", Global.ApplicationName, MessageBoxButton.OK, MessageBoxImage.Error);
                }
                Thread.Sleep(1000);
            }

            splashWindow.SetProgress("Loading Application");
            Thread.Sleep(1500);

            Global.blnmainWindow = false;

            // Create the main window, but on the UI thread.
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Invoker)delegate
            {
                if (Global.blnmainWindow != true) //Window already opened
                {
                    MainWindow mainWindow = new MainWindow();
                    mainWindow.Show();
                    Global.blnmainWindow=true;
                }
            });
        }

        public bool PerformDailyRecalc()
        {
            IniFile GlobalFile = new IniFile(LoadAppVariables.GlobalFile);

            GlobalFile.IniWriteValue("System", "DailyRecalc", DateTime.Now.ToString("dd/MM/yyyy"));

            // check how many Data populations we have and perform recalculate for each population

            try 
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                SqlCommand cmd = new SqlCommand("qryNewDailyPercentileRecalculation", con);
                cmd.CommandTimeout = Global.TimeOut;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();

                cmd.CommandText = "qryNewDailyPercentileRecalculation";
                cmd.CommandType = CommandType.StoredProcedure;
             
                cmd.ExecuteNonQuery();

                con.Close();
                cmd.Dispose();

                return true;
            }
            catch 
            {
                return false;
            }
        }

        public bool PerformAnnualUpdate()
        {
            IniFile GlobalFile = new IniFile(LoadAppVariables.GlobalFile);

            GlobalFile.IniWriteValue("System", "LastAnnualUpdate", DateTime.Now.Year.ToString());

            try
            {
                SqlConnection con = new SqlConnection(Global.ConnectionString);
                SqlCommand cmd = new SqlCommand("qryAnnualUpdate", con);
                cmd.CommandTimeout = Global.TimeOut;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();

                cmd.Parameters.Clear();
                cmd.CommandText = "qryAnnualUpdate";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@nYear", DateTime.Now.Year.ToString());

                cmd.ExecuteNonQuery();

                con.Close();
                cmd.Dispose();
                cmd.Parameters.Clear();

                return true;
            }
            catch
            {
                GlobalFile.IniWriteValue("System", "LastAnnualUpdate", Global.AnnualUpdate.ToString());

                return false;
            }
        }
    }
}
