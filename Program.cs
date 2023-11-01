using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Conditions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.Configuration;
using System.IO.Compression;
using System.Diagnostics.Eventing.Reader;

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
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string DBpath = ConfigurationManager.AppSettings["DBaddresspath"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
        static string logfilename = "";

        const int EXCELfile = 0;
        const int LOGFile = 1;
        const int ZIPFile = 2;
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
                DeleteSupportingFiles(appfolder, EXCELfile);
                DeleteSupportingFiles(appfolder, LOGFile);
                DeleteSupportingFiles(appfolder, ZIPFile);

                if (!Directory.Exists(appfolder))
                {
                    Directory.CreateDirectory(appfolder);
                    Directory.CreateDirectory(uploadfolder);
                    Directory.CreateDirectory(sharingfolder);
                }
                var config = new LoggerConfiguration();
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                Log.Logger = config.CreateLogger();

                Log.Information("Accurate Desktop ver.4 Automation -  by FAIRBANC");


                if (!OpenAppAndDBConfig())
                {
                    Log.Information("application automation failed !!");
                    return;
                }
            }
            catch (Exception ex) 
            { }
            finally { }
        }

        static bool OpenAppAndDBConfig()
        {
            return true;
        }
    }
}
    
