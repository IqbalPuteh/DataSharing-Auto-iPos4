using System;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using Serilog;
using System.Configuration;
using System.Diagnostics;
using FlaUI.Core.AutomationElements;

namespace iPos4DS_DTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static FlaUI.Core.Application appx;
        static Window AuAppMainWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
        static int step = 0;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string dbserveraddr = ConfigurationManager.AppSettings["dbserveraddress"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
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
                    Log.Information("Application automation failed when configure database (OpenAppAndDBConfig) !!!");
                    return;
                }
                if(!LoginApp())
                {
                    Log.Information("Application automation failed when login to app (loginApp) !!!");
                    return;
                }
            }
            catch (Exception ex)
            { Log.Information($"IPos automation error => {ex.ToString()}"); }
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
            var functionname = "OpenAppAndDBConfig";
            int steps = 0;
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
                    AuAppMainWindow = appx.GetMainWindow(automationUIA3);
                }
                catch { Log.Information($"[OpenAppAndDBConfig ]Error ketika mebuka mmnghandle iPos window process..."); }
                //* Wait until Accurate window ready
                Thread.Sleep(15000);
                //FlaUI.Core.Input.Wait.UntilResponsive(DesktopWindow.FindFirstChild(),TimeSpan.FromSeconds(4));

                //* Picking Koneksi Database main window
                var checkingele = "";
                var ParentEle = AuAppMainWindow.FindFirstChild(cf => cf.ByName("Koneksi Database"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                var ele = ParentEle.FindFirstChild(cf => cf.ByAutomationId("butServer", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                //ele.Click();
                //* check coordinates and try mouse click on the coordinates
                MouseClickaction(ele);

                //^ Traversing to 'lstData' descendant element from 'Koneksi Database' element
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("lstData", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //* Looking at 'lstData' items, and selecting server name base on item name
                ele = ele.FindFirstDescendant(cf => cf.ByName(dbserveraddr));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Click();

                ele = ParentEle.FindFirstDescendant(cf => cf.ByName("Pilih"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsButton().Focus();
                Thread.Sleep(1000);
                MouseClickaction(ele);
;
                ele = ParentEle.FindFirstChild(cf => cf.ByName("Cari Database"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsButton().Focus();
                Thread.Sleep(1000);
                MouseClickaction(ele);

                //* Traversing to 'lDatabase' element from 'Koneksi Database' element
                var listele = ParentEle.FindFirstChild(cf => cf.ByAutomationId("lDatabase")).AsListBox();
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                Thread.Sleep(1000); 
                listele.AsListBox().Items[0].Click();

                ele = ParentEle.FindFirstChild(cf => cf.ByName("Pilih"));
                ele.AsButton().Focus();
                Thread.Sleep(1000);
                MouseClickaction(ele);

                return true;
            }
            catch (Exception ex)
            {
                if (appx.ProcessId != null)
                {
                    appx.Close();
                }
                return false;
                Log.Information($"Error when executing {functionname} => {ex.Message}");
            }
        }

        private static string CheckingEle(AutomationElement? ele, int steps, string functionname)
        {
            var value = ele == null ? $"Automation error on #{steps} in function {functionname}..." : $"";
            return value;
        }

        private static bool LoginApp()
        {
            var functionname = "LoginApp";
            int steps = 0;
            try
            {
                //* Picking form login main window
                var checkingele = "";
                var ParentEle = AuAppMainWindow.FindFirstChild(cf => cf.ByName("Login"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                //tUser
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("tUser", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsTextBox().Enter(LoginId);

                //tPassword
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("tPassword", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsTextBox().Enter(LoginPassword);

                ele = ParentEle.FindFirstChild(cf => cf.ByName("Masuk"));
                ele.AsButton().Focus();
                Thread.Sleep(1000);
                MouseClickaction(ele);
                //ele.AsButton().Click();
                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.Message);
                return false;
            }
        }

        private static bool MouseClickaction(AutomationElement ele)
        {
            try
            {
                var elecornerpos = ele.GetClickablePoint();
                Mouse.MoveTo(elecornerpos.X, elecornerpos.Y);
                Mouse.Click();
                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing mouse click action on element {ele.AutomationId} => {ex.Message}");
                return false;
            }
        }

    }
}
    
