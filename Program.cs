using System;
using System.Diagnostics;
using System.Configuration;
using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Input;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.IO.Compression;

namespace iPos4DS_DTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static Application appx;
        //static Window AuAppMainWindow;
        static UIA3Automation automationUIA3;
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window;
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
        static string iposgudang = ConfigurationManager.AppSettings["namagudangdiipos"].ToUpper();
        static string shortcuttoipos = ConfigurationManager.AppSettings["shortcuttoipos"].ToUpper();
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
        static string logfilename = "";
        static int pid = 0;

        enum reportType
        {
            salesreport,
            arreport,
            masteroutletreport
        }

        [DllImport("user32.dll")]

        public static extern bool BlockInput(bool fBlockIt);
        
        private static AutomationElement WaitForElement(Func<AutomationElement> findElementFunc)
        {
            AutomationElement element = null;
            for (int i = 0; i < 2000; i++)
            {
                element = findElementFunc();
                if (element is not null)
                {
                    break;
                }

                //Thread.Sleep(1);
            }
            return element;
        }

        static bool IsFileExists(string path, string fileName)
        {
            string fullPath = Path.Combine(path, fileName);
            return File.Exists(fullPath);
        }

        static async Task Main(string[] args)
        {
            try
            {
                //* Call this method to disable keyboard input
#if DEBUG
                BlockInput(false);
#else
                BlockInput(true);
#endif
                var myFileUtil = new MyDirectoryManipulator();
                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                }
                var temp = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                Task.Run(() => Console.WriteLine($"[{ DateTime.Now.ToString("HH:mm:ss")} INF] {temp}")); 
                temp = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp}"));
                temp = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp}"));
                var config = new LoggerConfiguration();
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                Log.Logger = config.CreateLogger();

                Log.Information("iPOS ver.4 Automation - *** Started *** ");
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();

                if (!OpenAppAndDBConfig())
                {
                    Log.Information("application automation failed when running app (OpenAppAndDBConfig) !!!");
                    return;
                }
                if (!LoginApp())
                {
                    Log.Information("application automation failed when running app (LoginApp) !!!");
                    return;
                }
                if (!OpenReportParam("sales"))
                {
                    Log.Information("Application automation failed when running app (OpenReportParam) !!!");
                    return;
                }
                if (!SendingRptParam("sales"))
                {
                    Log.Information("Application automation failed when running app (SendingReportParam) !!!");
                    return;
                }

                if (!OpenReportParam("ar"))
                {
                    Log.Information("Application automation failed when running app (OpenReportParam) !!!");
                    return;
                }
                if (!SendingRptParam("ar"))
                {
                    Log.Information("Application automation failed when running app (SendingRptParam) !!!");
                    return;
                }
                if (!OpenReportParam("outlet"))
                {
                    Log.Information("Application automation failed when running app (OpenReportParam) !!!");
                    return;
                }
                if (!SendingRptParam("outlet"))
                {
                    Log.Information("Application automation failed when running app (SendingRptParam) !!!");
                    return;
                }
                if (await ZipandSendAsync() != true)
                {
                    Log.Information("Application automation failed when running app (ZipandSendAsync) !!!");
                    return;
                }
                if (!ClosingApp())
                {
                    Log.Information("application automation failed when running app (ClosingApp) !!!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"IPos automation error => {ex.ToString()}");
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] iPos automation error => {ex.ToString()}"));
            }
            finally
            {
                //* Call this method to enable keyboard input
                BlockInput(false);

                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] iPOS ver.4 Automation - ***   END   ***"));
                if (automationUIA3 is not null)
                {
                    automationUIA3.Dispose();
                }
                Log.CloseAndFlush();
            }
        }

        private static string CheckingEle(AutomationElement? ele, int steps, string functionname)
        {
            var value = ele is null ? $"Automation error on #{steps} in function {functionname}..." : $"";
            return value;
        }

        private static bool MouseClickaction(AutomationElement ele)
        {
            try
            {
                var elecornerpos = ele.GetClickablePoint();
                Mouse.MoveTo(elecornerpos.X + 2, elecornerpos.Y + 2);
                Mouse.Click();
                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing mouse click action on element {ele.AutomationId} => {ex.Message}");
                return false;
            }
        }

        static bool OpenAppAndDBConfig()
        {
            var functionname = "OpenAppAndDBConfig";
            int step = 0;
            try
            {

                // Specify the path to your shortcut
                string shortcutPath = $@"{shortcuttoipos}";
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = shortcutPath;
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = false;

                Process process = new Process();
                process.StartInfo = startInfo;
                automationUIA3 = new UIA3Automation();

                try
                {
                    appx = Application.Launch(process.StartInfo);
                    window = appx.GetMainWindow(automationUIA3);
                    pid = appx.ProcessId;
                    Thread.Sleep(30000);
                }
                catch { Log.Information($"[{functionname}] Error ketika mebuka mmnghandle iPos window process..."); return false; }

                //* Picking Koneksi Database main window
                var checkingele = "";
                var ParentEle = window.FindFirstDescendant(cf => cf.ByName("Koneksi Database"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                var ele = ParentEle.FindFirstChild(cf => cf.ByAutomationId("butServer", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
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
                listele.AsListBox().Items[0].Select();


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
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;

            }
        }

        private static bool LoginApp()
        {
            var functionname = "LoginApp";
            int step = 0;
            try
            {

                var checkingele = "";
                //* Picking form iPos 4 main windows
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                AutomationElement ParentEle = null;
                AutomationElement[] MainEle = window.FindAllChildren(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement elem in MainEle)
                {
                    if (elem.Properties.ProcessId != pid)
                    {
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        ParentEle = elem; break;
                    }
                }
                //* Picking form login main window
                ParentEle = ParentEle.FindFirstChild(cf => cf.ByName("Login"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                Thread.Sleep(1000);

                //tUser
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("tUser", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsTextBox().Enter(LoginId);
                Thread.Sleep(1000);

                //tPassword
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("tPassword", PropertyConditionFlags.None));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsTextBox().Enter(LoginPassword);
                Thread.Sleep(1000);

                ele = ParentEle.FindFirstChild(cf => cf.ByName("Masuk"));
                ele.AsButton().Focus();
                Thread.Sleep(1000);
                return MouseClickaction(ele);
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
        }

        private static bool OpenReportParam(string reportname)
        {
            var functionname = "OpenReportParam -> " + reportname;
            int step = 0;
            try
            {
                //* Picking form iPos 4 main windows
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                var checkingele = "";
                AutomationElement ParentEle = null;
                AutomationElement[] MainEle = window.FindAllChildren(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement elem in MainEle)
                {
                    if (!elem.Name.ToLower().Contains(LoginId.ToLower()))
                    {
                        Thread.Sleep(2000);
                    } else
                    {
                        ParentEle = elem; break;
                    }
                }
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //Ribbon Tabs
                ParentEle = ParentEle.FindFirstDescendant(cf => cf.ByName("The Ribbon"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                //Ribbon Tabs
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByName("Ribbon Tabs"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                MouseClickaction(ele);
                Thread.Sleep(500);

                //Penjualan
                ele = ele.FindFirstChild(cf => cf.ByName("Laporan"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                ele.AsTabItem().Select();
                ele.AsTabItem().Click();
                Thread.Sleep(1000);

                //Traversing to "Lower Ribbon" from Parent Element "The Ribbon"
                ele = ParentEle.FindFirstDescendant(cf => cf.ByName("Lower Ribbon"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                if (reportname == "sales")
                {
                    //'Penjualan' toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Penjualan"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }

                    //(This is) "Laporan Penjualan" toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Penjualan"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ParentEle.SetForeground();
                    MouseClickaction(ele);
                    Thread.Sleep(1000);

                    //(This is) "Laporan Penjualan" button
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Penjualan"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    MouseClickaction(ele);
                }
                else if (reportname == "ar")
                {
                    //'Hutang Piutang' toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Hutang Piutang"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }

                    //(This is) 'Laporan Piutang' toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Piutang"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ParentEle.SetForeground();
                    MouseClickaction(ele);
                    Thread.Sleep(1000);

                    //(This is) 'Laporan Pembayaran Piutang' button
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Pembayaran Piutang"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    MouseClickaction(ele);
                }
                else
                {
                    //'Master Data' toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Master"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }

                    //(This is) 'Laporan Master' toolbar
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Master"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ParentEle.SetForeground();
                    MouseClickaction(ele);
                    Thread.Sleep(1000);

                    //(This is) 'aftar Pelanggan' button
                    ele = ele.FindFirstDescendant(cf => cf.ByName("Daftar Pelanggan"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    MouseClickaction(ele);
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
        }

        private static bool SavingReport(string reportName)
        {
            var functionname = "SavingReport -> " + reportName;
            int step = 0;
            var checkingele = "";
            try
            {
                //* Picking report 'Preview' window,it's a direct child of main os window
                Thread.Sleep(5000);

                AutomationElement ParentEle;
                if (window == null || window is null)
                {
                    var au = new UIA3Automation();
                    ParentEle = window.FindFirstChild(cf => cf.ByAutomationId("PrintPreviewFormExBase"));
                }
                else
                { ParentEle = window.FindFirstChild(cf => cf.ByAutomationId("PrintPreviewFormExBase")); }

                for (int i = 1; i <= 120; i += 1) // ==> keep looking 'Preview' window for 10 minutes
                {
                    if (ParentEle != null)
                    {
                        break;
                    }
                    Thread.Sleep(5000);
                }
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //Dock Top
                var ele = ParentEle.FindFirstChild(cf => cf.ByName("Dock Top"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                //Main Menu
                ele = ele.FindFirstChild(cf => cf.ByName("Main Menu"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                Thread.Sleep(1000);

                //File
                ele = ele.FindFirstChild(cf => cf.ByName("File"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                MouseClickaction(ele);
                Thread.Sleep(1000);

                //then on the context menu (represent as list of  button element) travers to 
                //Export Document...
                ele = ele.FindFirstChild(cf => cf.ByName("Export Document..."));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                Mouse.MoveTo(ele.GetClickablePoint());
                Thread.Sleep(1000);

                //then from the 'Export Document...' button hovering action, move mouse to the new opened context menu 
                //XLSX File
                // ele = ele.FindFirstChild(cf => cf.ByName("XLSX File"));
                ele = ParentEle.Parent.FindFirstDescendant(cf => cf.ByName("XLSX File"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                //ele.SetForeground();
                MouseClickaction(ele);
                Thread.Sleep(5000);

                //when report parmeter windows with AutomationId: LinesForm show grab it
                ele = ParentEle.FindFirstChild(cf => cf.ByName("XLSX Export Options"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();

                //click 'OK' button
                ele = ele.FindFirstDescendant(cf => cf.ByName("OK"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                MouseClickaction(ele);
                Thread.Sleep(5000);

                //grabbing 'Save as' windows element
                ele = ParentEle.FindFirstChild(cf => cf.ByName("Save As"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();

                //1001
                var ele1 = ele.FindFirstDescendant(cf => cf.ByAutomationId("1001"));
                checkingele = CheckingEle(ele1, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele1.AsTextBox().Focus();
                
                var filename = reportName switch
                {
                    "sales" => "Sales_Data",
                    "ar" => "Repayment_Data",
                    _ => "Master_Outlet"
                };

                ele1.AsTextBox().Enter($@"{appfolder}\{filename}");

                //Save 
                ele1 = ele.FindFirstDescendant(cf => cf.ByName("Save"));
                checkingele = CheckingEle(ele1, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele1.Focus();
                MouseClickaction(ele1);

                //Function to check whether the file is already finnished created in intended folder
                DateTime startTime = DateTime.Now;
                Task.Delay(1000); 
                while (DateTime.Now - startTime < TimeSpan.FromMinutes(2))
                {
                    if (IsFileExists(appfolder, filename + ".xlsx"))
                    {
                        Log.Information($"File {filename}.xlsx saved successfully...");
                        break;
                    }
                    Task.Delay(5000);
                }
                if (!IsFileExists(appfolder, filename + ".xlsx"))
                {
                    Log.Information($"'Timeout' when saving {filename}.xlsx file...");
                    return false;
                }

                //Grabbbing 'Export' window
                ele = ParentEle.FindFirstChild(cf => cf.ByName("Export"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();

                //&No
                ele1 = ele.FindFirstDescendant(cf => cf.ByName("&No"));
                checkingele = CheckingEle(ele1, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele1.Focus();
                return MouseClickaction(ele1);
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
        }

        static bool SendingRptParam(string reportName)
        {
            var functionname = "SendingRptParam -> " + reportName;
            int step = 0;
            try
            {
                Thread.Sleep(5000);
                //* Picking iPos main window
                //var ParentEle = AuAppMainWindow.Parent.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));

                AutomationElement ParentEle;
                if (window == null || window is null)
                {
                    var au = new UIA3Automation();
                    ParentEle = window.FindFirstChild(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                }
                else
                { ParentEle = window.FindFirstChild(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring)); }
                for (int i = 1; i <= 120; i += 1) // ==> keep looking 'Preview' window for 10 minutes
                {
                    if (ParentEle != null)
                    {
                        break;
                    }
                    Thread.Sleep(5000);
                }

                var checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                //var eleName = "";
                //var imglstrptname = "";
                (var eleName, var imglstrptname) = reportName switch
                {
                    "sales" => ("frmLapPenjualan", "Laporan Penjualan Detail"),
                    "ar" => ("frmLapPiutangBayar", "Laporan Piutang - Pembayaran"),
                    _ => ("frmLapPelanggan", "" /*Not applicable here*/)
                };
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId(eleName));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                Thread.Sleep(1000);

                if (reportName != "outlet")
                {
                    if (reportName == "sales")
                    {
                        //Fill gudang value in 'cGudang' elements
                        ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("cGudang"));
                        checkingele = CheckingEle(ele, step += 1, functionname);
                        if (checkingele != "") { Log.Information(checkingele); return false; }

                        ele = ele.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit));
                        checkingele = CheckingEle(ele, step += 1, functionname);
                        if (checkingele != "") { Log.Information(checkingele); return false; }
                        ele.Focus();
                        ele.AsTextBox().Enter(iposgudang);
                        Thread.Sleep(1000);
                    }

                    BlockInput(false);

                    //dtTglDari
                    ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("dtTglDari"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ele.Focus();
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    ele.AsTextBox().Enter(DateManipul.GetFirstDate());
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    ele.AsTextBox().Enter(DateManipul.GetPrevMonth());
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    ele.AsTextBox().Enter(DateManipul.GetPrevYear());
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    if (reportName == "sales")
                    {
                        ele.AsTextBox().Enter("00");
                        Thread.Sleep(500);
                        Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                        ele.AsTextBox().Enter("00");
                        Thread.Sleep(500);
                    }

                    //dtTglSampai
                    ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("dtTglSampai"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ele.Focus();
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.LEFT);
                    ele.AsTextBox().Enter(DateManipul.GetLastDayOfPrevMonth());
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    ele.AsTextBox().Enter(DateManipul.GetPrevMonth());
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    ele.AsTextBox().Enter(DateManipul.GetPrevYear());
                    Thread.Sleep(500);
                    Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                    Thread.Sleep(500);
                    if (reportName == "sales")
                    {
                        ele.AsTextBox().Enter("23");
                        Thread.Sleep(500);
                        Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                        ele.AsTextBox().Enter("59"); ;
                        Thread.Sleep(500);
                    }
#if DEBUG
                    BlockInput(false);
#else
                    BlockInput(true);
#endif
                    //imgLst
                    ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("imgLst"));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ele.SetForeground();

                    //Laporan Penjualan Detail
                    ele = ele.FindFirstDescendant(cf => cf.ByName(imglstrptname));
                    checkingele = CheckingEle(ele, step += 1, functionname);
                    if (checkingele != "") { Log.Information(checkingele); return false; }
                    ele.Focus();
                    MouseClickaction(ele);
                    Thread.Sleep(500);
                }

                //butCetak
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("butCetak"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsButton().Click();
                //MouseClickaction(ele);
                Thread.Sleep(5000);


                //calling report save function
                if (!SavingReport(reportName))
                {
                    return false;
                }

                if (!ClosePreviewWindow())
                {
                    Log.Information("Application automation failed when running app (ClosePreviewWindow) !!!");
                    return false;
                }
                //butTutup
                var closeButtonEle = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("butTutup"));
                checkingele = CheckingEle(closeButtonEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                closeButtonEle.AsButton().Click();
                //MouseClickaction(closeButtonEle);
                Thread.Sleep(5000);

                return true;

            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
        }

        private static bool ClosePreviewWindow()
        {
            var functionname = "ClosePreviewWindow";
            int step = 0;
            var checkingele = "";
            try
            {
                //* Picking report 'Preview' window,it's a direct child of main os window
                Thread.Sleep(5000);

                AutomationElement ParentEle;
                if (window == null || window is null)
                {
                    var au = new UIA3Automation();
                    ParentEle = window.FindFirstChild(cf => cf.ByAutomationId("PrintPreviewFormExBase"));
                }
                else
                { ParentEle = window.FindFirstChild(cf => cf.ByAutomationId("PrintPreviewFormExBase")); }

                for (int i = 1; i <= 120; i += 1) // ==> keep looking 'Preview' window for 10 minutes
                {
                    if (ParentEle != null)
                    {
                        break;
                    }
                    Thread.Sleep(5000);
                }
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //Dock Top
                var ele = ParentEle.FindFirstChild(cf => cf.ByName("Dock Top"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();

                //Main Menu
                ele = ele.FindFirstChild(cf => cf.ByName("Main Menu"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                Thread.Sleep(1000);

                //File
                ele = ele.FindFirstChild(cf => cf.ByName("File"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                MouseClickaction(ele);
                Thread.Sleep(1000);

                //then on the context menu (represent as list of  button element) travers to 
                //Export Document...
                ele = ele.FindFirstChild(cf => cf.ByName("Exit"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();
                Mouse.MoveTo(ele.GetClickablePoint());
                MouseClickaction(ele);
                Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }

        }

        private static bool ClosingApp()
        {
            var functionname = "ClosingApp";
            int step = 0;
            var checkingele = "";
            try
            {
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                AutomationElement ParentEle = null;
                AutomationElement[] MainEle = window.FindAllChildren(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement elem in MainEle)
                {
                    if (!elem.Name.ToLower().Contains(LoginId.ToLower()))
                    {
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        ParentEle = elem; break;
                    }
                }
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //Automation Id: TitleBar
                var ele = ParentEle.FindFirstChild(cf => cf.ByAutomationId("TitleBar"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                Thread.Sleep(1000);

                //Automation Id: Close, as button
                ele = ele.FindFirstChild(cf => cf.ByAutomationId("Close"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.AsButton().Invoke();

                return true;
            } 
            catch (Exception ex) 
            {
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] Error during {functionname} function => {ex.Message}"));
                return false;
            }
        }

        private static async Task<bool> ZipandSendAsync()
        {
            try
            {
                Log.Information("Checking and deleting existing ZIP files...");
                var strDsPeriod = DateManipul.GetPrevYear() + DateManipul.GetPrevMonth();

                Log.Information("Moving standart excel reports file to uploaded folder...");
                // move excels files to Datafolder
                var path = appfolder + @"\Master_Outlet.xlsx";
                var path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_OUTLET.xlsx";
                File.Move(path, path2, true);
                path = appfolder + @"\Sales_Data.xlsx";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_SALES.xlsx";
                File.Move(path, path2, true);
                path = appfolder + @"\Repayment_Data.xlsx";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_AR.xlsx";
                File.Move(path, path2, true);

                // set zipping name for files
                Log.Information("Zipping Transaction file(s)");
                var strZipFile = dtID + "-" + dtName + "_" + strDsPeriod + ".zip";
                ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the ZIP file to the API server 
                using (var mycUrl = new cUrlClass('Y', issandbox.ToArray().First(), "", sharingfolder + Path.DirectorySeparatorChar + strZipFile))
                {
                    Log.Information("Sending ZIP file to the API server...");
                    var strStatusCode = "0"; // varible for debugging Curl test
                    strStatusCode = await mycUrl.SendRequestAsync();
                    Thread.Sleep(5000);
                    if (strStatusCode == "200")
                    {
                        Log.Information("DATA TRANSACTION SHARING - SELESAI");
                    }
                    else
                    {
                        Log.Information("Failed to send TRANSACTION file to API server... => " + strStatusCode);
                    }

                }
                Log.CloseAndFlush();
                using (var mycUrl = new cUrlClass('Y', issandbox.ToArray().First(), "", appfolder + Path.DirectorySeparatorChar + logfilename))
                {
                    Log.Information("Sending log file to the API server...");
                    Task.Run(() => Console.WriteLine("Sending log file to the API server..."));
                    var strStatusCode = "0"; // varible for debugging Curl test
                    strStatusCode = await mycUrl.SendRequestAsync();
                    Thread.Sleep(10000);
                    if (strStatusCode != "200")
                    {
                        throw new Exception ("Failed to send LOG file to API server...");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] Error during ZIP and cUrl send => {ex.Message}"));
                return false;
            }
        }
    }

}
    
