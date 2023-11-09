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
using static System.Windows.Forms.Design.AxImporter;

namespace iPos4DS_DTTest // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static Application appx;
        static Window AuAppMainWindow;
        static UIA3Automation automationUIA3 = new UIA3Automation();
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window = automationUIA3.GetDesktop();
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
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
        static string logfilename = "";

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
                //* Call this method to disable keyboard input
                BlockInput(true);

                var myFileUtil = new MyDirectoryManipulator();
                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                }
                Console.WriteLine(myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel));
                Console.WriteLine(myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log));
                Console.WriteLine(myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip));


                var config = new LoggerConfiguration();
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                Log.Logger = config.CreateLogger();

                Log.Information("iPOS ver.4 Automation - by FAIRBANC *** Started! *** ");


                //if (!OpenAppAndDBConfig())
                //{
                //    Log.Information("application automation failed when running app (openappanddbconfig) !!!");
                //    return;
                //}
                //if (!LoginApp())
                //{
                //    Log.Information("application automation failed when running app (loginapp) !!!");
                //    return;
                //}
                if (!OpenSalesReport())
                {
                    Log.Information("Application automation failed when running app (OpenSalesReport) !!!");
                    return;
                }
                if(!SendingReportParam())
                {
                    Log.Information("Application automation failed when running app (SendingReportParam) !!!");
                    return;
                }
            }
            catch (Exception ex)
            { Log.Information($"IPos automation error => {ex.ToString()}"); }
            finally
            {
                //* Call this method to enable keyboard input
                BlockInput(false);

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
            int step = 0;
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
                catch { Log.Information($"[{functionname}] Error ketika mebuka mmnghandle iPos window process..."); }
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
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;

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
            int step = 0;
            try
            {
                //* Picking form login main window
                var checkingele = "";

                var ParentEle = window.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
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
                ele.AsButton().Invoke();
                //MouseClickaction(ele);

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
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

        private static bool OpenSalesReport()
        {
            var functionname = "OpenSalesReport";
            int step = 0;
            try
            {
                //* Picking form iPos 4 main windows
                var checkingele = "";
                var ParentEle = window.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                while (!ParentEle.Name.ToLower().Contains(LoginId.ToLower()))
                {
                    Thread.Sleep(2000);
                }

                //Ribbon Tabs
                ParentEle = ParentEle.FindFirstDescendant(cf => cf.ByName("The Ribbon"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                //ParentEle.SetForeground();

                //Ribbon Tabs
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByName("Ribbon Tabs"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                Thread.Sleep(500);

                //Penjualan
                ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                MouseClickaction(ele);
                Thread.Sleep(1000);

                //Traversing to "Lower Ribbon" from Parent Element "The Ribbon"
                ele = ParentEle.FindFirstDescendant(cf => cf.ByName("Lower Ribbon"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                //Penjualan toolbar
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

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
        }

        static bool SendingReportParam()
        {
            var functionname = "SendingReportParam";
            int step = 0;
            try
            {
                Thread.Sleep(5000);
                //* Picking iPos main window
                //var ParentEle = AuAppMainWindow.Parent.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));

                AutomationElement ParentEle;
                if (AuAppMainWindow == null || AuAppMainWindow is null)
                {
                    var au = new UIA3Automation();
                    ParentEle = window.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring));
                }
                else
                { ParentEle = AuAppMainWindow.Parent.FindFirstDescendant(cf => cf.ByName("i P o s", PropertyConditionFlags.MatchSubstring)); }
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

                ParentEle = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("frmLapPenjualan"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ParentEle.SetForeground();
                Thread.Sleep(1000);

                //cGudang
                var ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("cGudang"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }

                ele = ele.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                ele.AsTextBox().Enter(iposgudang);
                Thread.Sleep(1000);

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
                ele.AsTextBox().Enter("01");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("01");
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("2023");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("00");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                ele.AsTextBox().Enter("00");
                Thread.Sleep(500);

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
                ele.AsTextBox().Enter("31");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("12");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("2023");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                Thread.Sleep(500);
                ele.AsTextBox().Enter("23");
                Thread.Sleep(500);
                Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.RIGHT);
                ele.AsTextBox().Enter("59"); ;
                Thread.Sleep(500);

                //imgLst
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("imgLst"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.SetForeground();

                //Laporan Penjualan Detail
                ele = ele.FindFirstDescendant(cf => cf.ByName("Laporan Penjualan Detail"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                MouseClickaction(ele);
                Thread.Sleep(500);

                //butCetak
                ele = ParentEle.FindFirstDescendant(cf => cf.ByAutomationId("butCetak"));
                checkingele = CheckingEle(ParentEle, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele.Focus();
                MouseClickaction(ele);
                Thread.Sleep(5000);

                //calling report save function
                if(!SavingReport01())
                {
                    return false;
                }
                return true;


            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
            finally
            {
                BlockInput(true);
            }
        }

        static bool SavingReport01()
        {
            var functionname = "SavingReport01";
            int step = 0;
            try
            {

                //* Picking report 'Preview' window,it's a direct child of main os window
                Thread.Sleep(5000);

                AutomationElement ParentEle;
                if (AuAppMainWindow == null || AuAppMainWindow is null)
                {
                    var au = new UIA3Automation();
                    ParentEle = window.FindFirstDescendant(cf => cf.ByAutomationId("PrintPreviewFormExBase"));
                }
                else
                { ParentEle = AuAppMainWindow.Parent.FindFirstDescendant(cf => cf.ByAutomationId("PrintPreviewFormExBase")); }

                for (int i = 1; i <= 120; i += 1) // ==> keep looking 'Preview' window for 10 minutes
                {
                    if (ParentEle != null) 
                    {
                        break; 
                    } 
                    Thread.Sleep(5000); 
                }

                //Dock Top
                var ele = ParentEle.FindFirstChild(cf => cf.ByName("Dock Top"));
                var checkingele = CheckingEle(ele, step += 1, functionname);
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
                Mouse.MoveTo (ele.GetClickablePoint());
                Thread.Sleep(1000);

                //then from the 'Export Document...' button hovering action, move mouse to the new opened context menu 
                //XLSX File
                // ele = ele.FindFirstChild(cf => cf.ByName("XLSX File"));
                ele = ParentEle.Parent.FindFirstDescendant(cf => cf.ByName("XLSX File"));
                checkingele = CheckingEle(ele, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                //ele.SetForeground();
                Mouse.MoveTo(ele.GetClickablePoint());
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
                ele1.AsTextBox().Enter($@"{appfolder}\sales");

                //Save 
                ele1 = ele.FindFirstDescendant(cf => cf.ByName("Save"));
                checkingele = CheckingEle(ele1, step += 1, functionname);
                if (checkingele != "") { Log.Information(checkingele); return false; }
                ele1.Focus();
                MouseClickaction (ele1);
                Thread.Sleep(5000);

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
                MouseClickaction(ele1);

                return closePreviwWindow(ParentEle, step, functionname);


            }
            catch (Exception ex)
            {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }

        }
        private static bool closePreviwWindow(AutomationElement ParentEle, int step, string functionname)
        {
            try 
            {
                //Dock Top
                var ele = ParentEle.FindFirstChild(cf => cf.ByName("Dock Top"));
                var checkingele = CheckingEle(ele, step += 1, functionname);
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
            catch (Exception ex) {
                Log.Information($"Error when executing {functionname} => {ex.Message}");
                return false;
            }
            
        }
    }

}
    
