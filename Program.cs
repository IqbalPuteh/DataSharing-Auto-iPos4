using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Conditions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.Configuration;
using System.Diagnostics;
using Microsoft.Win32.TaskScheduler;
using System.Text;

namespace iPos4DS_DTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static Application appx;
        static Window DesktopWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static int step = 0;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string DBpath = ConfigurationManager.AppSettings["DBaddresspath"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string isrunbyscheduler = ConfigurationManager.AppSettings["isrunbywindowsscheduler"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
        static string logfilename = "";

        const UInt32 WM_CLOSE = 0x0010;

        const int SALESREPORT = 0;
        const int ARREPORT = 0;
        const int OUTLETREPORT = 1;

        private static AutomationElement WaitForElement(Func<AutomationElement> findElementFunc)
        {
            AutomationElement element = null;
            for (int i = 0; i < 2000; i++)
            {
                element = findElementFunc();
                if (element != null)
                {
                    break;
                }

                Thread.Sleep(1);
            }
            return element;
        }

        static void Main(string[] args)
        {
            try 
            {
                var myFileUtil = new MyDirectoryManipulator();
                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                }
                myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);


                var config = new LoggerConfiguration();
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                Log.Logger = config.CreateLogger();

                Log.Information("iPOS ver.4 Automation - by FAIRBANC *** Started! *** ");


                if (!OpenAppAndDBConfig())
                {
                    Log.Information("Application automation failed !!");
                    return;
                }
            }
            catch (Exception ex) 
            { Log.Information($"IPos automation error => {ex.ToString()}");  }
            finally 
            {
                Log.Information("iPOS ver.4 Automation - *** END ***");
                if (automationUIA3 != null)
                {
                    automationUIA3.Dispose();
                }
                Log.CloseAndFlush();
            }
        }

        static bool OpenAppAndDBConfig()
        {
            try
            {
                // Specify the path to your shortcut
                string shortcutPath = @"C:\Users\iputeh\Desktop\iPos 4.0 Program Toko.lnk";
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = shortcutPath;
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = false;

                Process process = new Process();
                process.StartInfo = startInfo;
                //process.Start();

                Thread.Sleep(15000);

                try
                {
                    appx = Application.Launch(process.StartInfo);
                    DesktopWindow = appx.GetMainWindow(automationUIA3);
                }
                catch { }
                //* Wait until Accurate window ready
                Thread.Sleep(15000);
                
                //FlaUI.Core.Input.Wait.UntilResponsive(DesktopWindow.FindFirstChild(),TimeSpan.FromSeconds(4));
                //appx.Close();

                return true;
            }
            catch
            {
                if (appx.ProcessId != null)
                {
                    appx.Close();
                }
                return false;
            }
        }

        private static bool RunIPOSByScheduler()
        {
            try
            {
                // Create a scheduled task to run your application with elevated privileges
                using (TaskService ts = new TaskService())
                {
                    TaskDefinition td = ts.NewTask();
                    td.RegistrationInfo.Description = "FAIRBANC - Run iPos application for automation using FlaUI";

                    td.Principal.RunLevel = TaskRunLevel.Highest; // Run with highest privileges

                    td.Triggers.Add(new TimeTrigger { StartBoundary = DateTime.Now + TimeSpan.FromSeconds(10) }); // Specify the start time of the task

                    td.Actions.Add(new ExecAction($"{appExe}", null, null)); // Specify the path to your application

                    ts.RootFolder.RegisterTaskDefinition("FAIRBANC - Run iPos application for automation", td); // Register the task in the root folder
                }

                // Wait for the task to start
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(15));

                // Attach FlaUI to the running process
                appx = Application.AttachOrLaunch(new ProcessStartInfo($"{appExe}"));

                return true;
            }
            catch { 
                return false;
                throw;
            }
        }
    }
}
    
