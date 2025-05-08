using HtmlAgilityPack;

using LiteDB;

using Microsoft.Extensions.Hosting;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace PrevailETL
{
    class Program
    {        
        public static List<ControlPrevail> ControlPrevailList = new List<ControlPrevail>();
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        public static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // When you don't want the ProcessId, use this overload and pass 
        // IntPtr.Zero for the second parameter
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        /// The GetForegroundWindow function returns a handle to the 
        /// foreground window.
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BringWindowToTop(HandleRef hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        //one source
        public static int SW_HIDE = 0;
        public static int SW_SHOWNORMAL = 1;
        public static int SW_SHOWMINIMIZED = 2;
        public static int SW_SHOWMAXIMIZED = 3;
        public static int SW_SHOWNOACTIVATE = 4;
        public static int SW_RESTORE = 9;
        public static int SW_SHOWDEFAULT = 10;

        //other source
        public static int SW_SHOW = 5;

        public static void ForceForegroundWindow(IntPtr hWnd)
        {
            uint foreThread = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
            uint appThread = GetCurrentThreadId();
            const uint SW_SHOW = 5;

            if (foreThread != appThread)
            {
                AttachThreadInput(foreThread, appThread, true);
                BringWindowToTop(hWnd);
                ShowWindow(hWnd, SW_SHOW);
                AttachThreadInput(foreThread, appThread, false);
            }
            else
            {
                BringWindowToTop(hWnd);
                ShowWindow(hWnd, SW_SHOW);
            }
        }
        static void Main(string[] args)
        {
            Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "--auto-open-devtools-for-tabs --remote-debugging-port=9222 --user-data-dir=\"C:\\Users\\Luciano\\AutomationData\"");
            ChromeOptions options = new ChromeOptions();
            options.DebuggerAddress = "127.0.0.1:9222";

            var driver = new ChromeDriver(options);
            Console.WriteLine(driver.SessionId.ToString());

            //BackProcessCsvData("all");
            //BackProcessCsvData("ae_es_view");
            //BackProcessCsvData("ae_ld_view");
            //BackProcessCsvData("ae_dr_view");
            // Solo activar en caso de procesar datos anteriores
            //ProcessJsonData();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            while (true)
            {
                try
                {
                    CheckAlert(driver);

                    driver.ExecuteScript("document.alert = window.alert = alert = () => {};", null);
                    driver.ExecuteCdpCommand("Page.addScriptToEvaluateOnNewDocument", new Dictionary<string, object>() {
                        { "source", "alert = function() {}"} 
                    });
                    var currentTime = sw.Elapsed;
                    if (currentTime >= new TimeSpan(0, 15, 0) && DateTime.Now.Hour != 7)
                    {                        
                        throw new Exception("Ellapsed 15 Minutes");
                    }
                    

                    CheckAlert(driver);

                    try
                    {
                        ProcessOperationalInfo(driver);
                        ProcessDrillInfo(driver);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"No se pudo procesar drills, excepción: {e.Message}");
                    }

                   
                    driver.Url = "https://defaulttrends.joyglobal.com";
                    try
                    {
                        var uName = driver.FindElement(By.Name("username"));
                        if (uName != null)
                        {
                            var pass = driver.FindElement(By.Name("password"));                         
                            var btnSubmit = driver.FindElement(By.ClassName("btn"));
                            btnSubmit.Click();
                            Thread.Sleep(15000);
                        }
                    }catch(Exception e)
                    {
                        Console.WriteLine($"Ya conectado");
                    }
                    Console.WriteLine($"{DateTime.Now.ToString("s")}\tVentana Actual: {driver.Title}");
                    if (driver.Title.Contains("Grafana"))
                    {
                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Datos LD2350-2209");
                        IterateLDDashboard(driver, "LD-2350-2209");
                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Datos LD2350-2223");
                        IterateLDDashboard(driver, "LD-2350-2223");
                        
                        TransferPrevailAlert(ControlPrevailList);

                        ControlPrevailList = new List<ControlPrevail>();
                        driver.Navigate().Refresh();
                    }
                    Console.WriteLine(driver.Title);
                    
                    driver.Url = "https://analytics.joyglobal.com/AEViewer/AELive.html";
                    CheckAlert(driver);
                    driver.Url = "https://analytics.joyglobal.com/AEViewer/AELive.html";
                    if (driver.Title.Contains("Live"))
                    {
                        driver.Url = "https://analytics.joyglobal.com/AEViewer/AELive.html";                        
                        

                        //JSON VERSION
                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Alarmas Palas");
                        ProcessEquipmentJsonData(driver, "ae_es_alarmonly_view");

                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Datos ES");
                        ProcessEquipmentJsonData(driver, "ae_es_view");

                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Datos LD");
                        ProcessEquipmentJsonData(driver, "ae_ld_view");

                        Console.WriteLine($"{DateTime.Now.ToString("s")}\tBuscando Datos DR");
                        ProcessEquipmentJsonData(driver, "ae_dr_view");

                        ProcessJsonData();
                        if (ControlPrevailList.Count > 0)
                        {
                            Console.WriteLine($"{DateTime.Now.ToString("s")}\tIngresando resultados a control de prevail");
                            TransferPrevailAlert(ControlPrevailList);
                        }
                        ControlPrevailList = new List<ControlPrevail>();
                        driver.Navigate().Refresh();
                    }
                }
                catch (Exception e)
                {
                    Thread.Sleep(5000);
                    Console.WriteLine($"Handle muerto, {e.Message}, {e.StackTrace}");
                    Process[] chromeInstances = Process.GetProcessesByName("chrome");

                    foreach (Process p in chromeInstances)
                        p.Kill();

                    Thread.Sleep(15000);


                    if (e.Message.Contains("Grafana closed the connection")
                        || e.Message.Contains("#GotoDefaultTrending") 
                        || e.Message.Contains("Ellapsed 15 Minutes"))
                    {
                        string appPath = @"C:\Users\Administrator\source\repos\PrevailETL2\PrevailETL\bin\Debug\netcoreapp3.1\PrevailETL.exe";
                        ProcessStartInfo Info = new ProcessStartInfo();
                        Info.Arguments = "/C ping 127.0.0.1 -n 10 && \"" + appPath + "\"";
                        Info.WindowStyle = ProcessWindowStyle.Maximized;
                        Info.FileName = "cmd.exe";
                        Process.Start(Info);
                        Environment.Exit(0);
                    }

                    else
                    {
                        Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "--remote-debugging-port=9222 --user-data-dir=\"C:\\Users\\Carlos\\AutomationData\"");
                        Thread.Sleep(15000);                    
                        options.DebuggerAddress = "127.0.0.1:9222";
                        driver = new ChromeDriver(options);
                        Console.WriteLine(driver.SessionId.ToString());
                    }
                    
                    
                }
                Console.WriteLine($"{DateTime.Now.ToString("s")}\tEsperando 60s");
                Thread.Sleep(60000);

            }
        }

        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        public static void CheckAlert(ChromeDriver driver)
        {
            try
            {
                while (true)
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                    var alert = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());
                    if (alert != null)
                        alert.Accept();
                }
            }
            catch (Exception e)
            {
                //exception handling
                Console.WriteLine($"No alerts");
            }
        }

        private static void ProcessDrillInfo(ChromeDriver driver)
        {
            driver.Navigate().GoToUrl("https://analytics.joyglobal.com/Prevail/Drill/Production/HoleSummary.aspx");
            Thread.Sleep(5000);
            driver.Navigate().GoToUrl("https://analytics.joyglobal.com/Prevail/Drill/Production/HoleSummary.aspx");

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("$(\".timeRangeSelectDropDownArrow\")[1].click();");
            Thread.Sleep(5000);

            js.ExecuteScript("$(\"#ctl00_radioButtonListDateType_4\").click();");
            Thread.Sleep(5000);

            js.ExecuteScript("$(\"#ctl00_cmdSelectMachines\").click();");
            Thread.Sleep(15000);

            js.ExecuteScript("$(\"#ctl00_contentPlaceHolderMain_roundedPanelHoleSummaryGrid_imageBtnHoleSummaryGrid\").click();");
            Thread.Sleep(5000);

            // We open the HoleSummary file and process every hole
            HtmlDocument htmlDoc = new HtmlDocument();
            string file = @"C:\Users\Administrator\Downloads\HoleSummary.xls";
            htmlDoc.LoadHtml(File.ReadAllText(file));

            List<Dictionary<string, string>> drillList = new List<Dictionary<string, string>>();
            foreach (HtmlNode table in htmlDoc.DocumentNode.SelectNodes("//table"))
            {
                List<string> fields = new List<string>();
                foreach (HtmlNode header in table.SelectNodes("thead"))
                {
                    HtmlNode row = header.SelectNodes("tr")[0];
                    foreach (HtmlNode cell in row.SelectNodes("th"))
                    {
                        if (cell.InnerText != " ")
                        {
                            fields.Add(cell.InnerText.Replace(" ", "").Replace("(", "").Replace(")", ""));
                        }
                        else
                        {
                            fields.Add(cell.InnerText);
                        }
                    }
                }
                foreach (HtmlNode header in table.SelectNodes("tbody"))
                {
                    if (header.SelectNodes("tr") != null)
                    {
                        foreach (HtmlNode row in header.SelectNodes("tr"))
                        {
                            int cIndex = 0;
                            Dictionary<string, string> aData = new Dictionary<string, string>();
                            foreach (HtmlNode cell in row.SelectNodes("td"))
                            {
                                if (fields[cIndex] != "StartTime" && fields[cIndex] != "EndTime")
                                {
                                    aData[fields[cIndex]] = cell.InnerText;
                                }
                                else
                                {
                                    aData[fields[cIndex]] = DateTime.Parse(cell.InnerText, null, System.Globalization.DateTimeStyles.AssumeLocal).ToUniversalTime().ToString("s");
                                }
                                cIndex++;
                            }
                            drillList.Add(aData);
                        }
                    }
                }
            }

            FileInfo f = new FileInfo(file);
            string movedName = @$"c:\PrevailDrill\processed_{Path.GetFileName(f.FullName)}";
            if (File.Exists(movedName))
            {
                File.Delete(movedName);
            }
            File.Move(file, movedName);

            // For every hole, we load it's detail page and proccess the detailed information
            List<Dictionary<string, string>> finalDrillList = new List<Dictionary<string, string>>();
            foreach (var baseDrill in drillList)
            {
                try
                {
                    if (DateTime.Parse(baseDrill["StartTime"], null, System.Globalization.DateTimeStyles.AssumeUniversal) > DateTime.UtcNow.AddHours(-12))
                    {
                        Console.WriteLine($"Pozo {baseDrill["HoleId"]} aun no cumple 12 Horas.");
                        continue;
                    }
                    string hid = $"{baseDrill["Equipment"]}_{baseDrill["HoleId"]}_{DateTime.Parse(baseDrill["StartTime"]).Ticks}";
                    string fix = @$"C:\PrevailDrill\DrillIndex\dindex_{hid}.txt";

                    if (File.Exists(fix)) 
                        continue;

                    driver.Navigate().GoToUrl($"https://analytics.joyglobal.com/Prevail/Drill/Production/HoleDetailActual.aspx?ID={baseDrill["HoleId"]}");

                    js.ExecuteScript("$(\"#ctl00_contentPlaceHolderMain_roundedPanelHoleDetail_imageBtnHoleDetailActualChart\").click();");
                    Thread.Sleep(5000);
                    
                    var htmlDetail = new HtmlDocument();
                    string fileDrill = @"C:\Users\Administrator\Downloads\HoleDetail.xls";
                    htmlDetail.LoadHtml(File.ReadAllText(fileDrill));

                    FileInfo fd = new FileInfo(fileDrill);
                    string movedDrill = @$"c:\PrevailDrill\processed_{Path.GetFileName(fd.FullName)}";
                    bool procesado = false;


                    try
                    {
                        ICollection<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
                        foreach (HtmlNode table in htmlDetail.DocumentNode.SelectNodes("//table"))
                        {
                            List<string> fields = new List<string>();
                            foreach (HtmlNode header in table.SelectNodes("thead"))
                            {
                                foreach (HtmlNode row in header.SelectNodes("tr"))
                                {
                                    foreach (HtmlNode cell in row.SelectNodes("th"))
                                    {
                                        fields.Add(cell.InnerText.Replace(" ", "").Replace("(", "").Replace(")", ""));
                                    }
                                }
                            }
                            foreach (HtmlNode header in table.SelectNodes("tbody"))
                            {
                                foreach (HtmlNode row in header.SelectNodes("tr"))
                                {
                                    int cIndex = 0;
                                    Dictionary<string, string> aData = new Dictionary<string, string>();
                                    foreach (HtmlNode cell in row.SelectNodes("td"))
                                    {
                                        aData[fields[cIndex]] = cell.InnerText;
                                        cIndex++;
                                    }
                                    data.Add(aData);
                                }
                            }
                        }

                        baseDrill.Remove("&nbsp;");
                        baseDrill.Remove(" ");
                        baseDrill.Remove("OperatorId");
                        baseDrill["details"] = JsonConvert.SerializeObject(data);
                        finalDrillList.Add(baseDrill);

                        var fixf = File.Create(fix);
                        fixf.Close();
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"Invalid drill info, {e.Message}");
                    }
                    

                   
                    if (File.Exists(movedDrill))
                    {
                        File.Delete(movedDrill);
                    }
                    File.Move(fileDrill, movedDrill);
                    Thread.Sleep(2000);
                    procesado = true;
                }
                catch(Exception e)
                {
                    Console.WriteLine($"An error ocurred while processing drill");
                }
            }

            // At the end, we generate a composite package and transmit it to an InputStream.
            var success = TransferDrillData(finalDrillList, "");
            if (success)
            {
                Console.WriteLine($"Transfered {drillList.Count}");
            }
            else
            {
                Console.WriteLine($"Failed to Transfer {drillList.Count}");
            }
            driver.Navigate().GoToUrl("https://analytics.joyglobal.com/Prevail/Drill/Production/HoleSummary.aspx");
        }

        private static void ClearDirectory(string[] folders)
        {
            foreach (string folder in folders)
            {
                DirectoryInfo fi = new DirectoryInfo(folder);
                if (fi.CreationTime < DateTime.Now.AddDays(-5))
                    Directory.Delete(folder, true);
            }
        }

        private static void ClearFolders()
        {
            string[] aeFolders = Directory.GetDirectories(@"C:\PrevailIndex\AE");
            ClearDirectory(aeFolders);
           
            string[] drillFolder = Directory.GetDirectories(@"C:\PrevailIndex\Grafana\Drills");
            ClearDirectory(drillFolder);

            string[] loaderFolders = Directory.GetDirectories(@"C:\PrevailIndex\Grafana\Loaders");
            ClearDirectory(loaderFolders);

            string[] grafanaFolders = Directory.GetDirectories(@"C:\PrevailIndex\Grafana\Sholvers");
            ClearDirectory(grafanaFolders);
        }

        private static void OpenNewWindow(string url, ChromeDriver driver)
        {            
            driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.Navigate().GoToUrl(url);            
        }

        public static bool CheckWindow(Expression<Func<IWebDriver, bool>> predicateExp, ChromeDriver driver)
        {
            var predicate = predicateExp.Compile();
            CheckAlert(driver);
            foreach (var handle in driver.WindowHandles)
            {
                driver.SwitchTo().Window(handle);
                Thread.Sleep(2000);
                if (predicate(driver))
                {
                    return true;
                }
            }

            Console.WriteLine(string.Format("Unable to find window with condition: '{0}'", predicateExp.Body));
            return false;
        }

        public static void SwitchToWindow(Expression<Func<IWebDriver, bool>> predicateExp, ChromeDriver driver)
        {
            var predicate = predicateExp.Compile();
            foreach (var handle in driver.WindowHandles)
            {
                driver.SwitchTo().Window(handle);
                Thread.Sleep(2000);
                if (predicate(driver))
                {
                    return;
                }
            }

            Console.WriteLine(string.Format("Unable to find window with condition: '{0}'", predicateExp.Body));
        }

        public static bool IterateESDashboard(ChromeDriver driver, string equipment)
        {
            ICollection<GrafanaQuery> queries = new List<GrafanaQuery>()
            {
                  new GrafanaQuery()
                {
                    Alias = "ES_Op_Dipper_Button_Trip_Status",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Op_Dipper_Button_Trip_Status",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Sys_RHM_SAO_State",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Sys_RHM_SAO_State",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "2m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Opt_PayLoad_SAO_State",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Opt_PayLoad_SAO_State",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "2m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_Mtr_Ft_Main_Volt_V",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_Mtr_Ft_Main_Volt_V",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "15m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Sys_Run_hr",
                    Query = new {
                        aggregator = "max",
                        metric = "ES_Sys_Run_hr",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Air_Main_Pressure_PSI",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Air_Main_Pressure_PSI",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "2m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_Mtr_Ft_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_Mtr_Ft_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_Mtr_Ft_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_Mtr_Ft_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_Mtr_Rr_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_Mtr_Rr_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_Mtr_Rr_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_Mtr_Rr_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_Mtr_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_Mtr_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_Mtr_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_Mtr_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_RtF_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_RtF_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_RtF_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_RtF_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_Rr_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_Rr_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_Rr_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_Rr_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_LtF_DE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_LtF_DE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_Mtr_LtF_NDE_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_Mtr_LtF_NDE_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Ft_LftBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Ft_LftBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Ft_RtBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Ft_RtBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Ft_LftBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Ft_LftBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Ft_RtBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Ft_RtBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Rr_LftBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Rr_LftBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Rr_RtBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Rr_RtBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Rr_LftBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Rr_LftBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Hst_GB_Mtr_Rr_RtBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Hst_GB_Mtr_Rr_RtBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_LftBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_LftBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_RtBrgRed2_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_RtBrgRed2_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_LftBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_LftBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_RtBrgRed1_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_RtBrgRed1_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_LftBrgSS_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_LftBrgSS_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Cwd_GB_RtBrgSS_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Cwd_GB_RtBrgSS_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_RtF_Lube_Smp_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_RtF_Lube_Smp_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_RtF_SwgShftBrg_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_RtF_SwgShftBrg_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_Rr_Lube_Smp_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_Rr_Lube_Smp_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_Rr_SwgShftBrg_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_Rr_SwgShftBrg_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_LtF_Lube_Smp_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_LtF_Lube_Smp_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Swg_GB_LtF_SwgShftBrg_Temp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Swg_GB_LtF_SwgShftBrg_Temp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "ES_Pwr_ISU_Main_Volt_V",
                    Query = new {
                        aggregator = "sum",
                        metric = "ES_Pwr_ISU_Main_Volt_V",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                }
            };

            List<GrafanaMeasurement> measurements = new List<GrafanaMeasurement>();
            DateTime start = DateTime.UtcNow.AddHours(-3.2);
            foreach (GrafanaQuery query in queries)
            {
                measurements.AddRange(GetTrendData(driver, query, start, @$"C:\PrevailIndex\Grafana\Sholvers"));
            }
            
            if (measurements.Count > 0)
            {
                bool insertResult = TransferGrafanaData(measurements, "3cf0f328c7e938a900500d8349c37392b42630a7");
                if (insertResult)
                {
                    Console.WriteLine($"Sent {measurements.Count}");
                }
            }
            else
            {
                Console.WriteLine("No new data available on Grafana");
            }
            return true;
        }

        private static void ProcessOperationalInfo(ChromeDriver driver)
        {
            List<string> flotaButton = new List<string>()
            {
                "Drill", "Loader", "Shovel"
            };

            Dictionary<string, string> fileDict = new Dictionary<string, string>()
            {
                {"Drill", "OperatingPractices" },
                {"Loader", "OperatingPracticesDetail" },
                {"Shovel", "OperatingPractices" },
            };


            List<Dictionary<string, object>> failureList = new List<Dictionary<string, object>>();

            foreach (string flota in flotaButton)
            {
                try
                {
                    driver.Navigate().GoToUrl($"https://analytics.joyglobal.com/Prevail/{flota}/Maintenance/OperatingPractices.aspx");                    
                    Thread.Sleep(5000);
                    driver.Navigate().GoToUrl($"https://analytics.joyglobal.com/Prevail/{flota}/Maintenance/OperatingPractices.aspx");
                    Thread.Sleep(5000);

                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("$(\".timeRangeSelectDropDownArrow\")[1].click();");
                    Thread.Sleep(5000);

                    js.ExecuteScript("$(\"#ctl00_radioButtonListDateType_4\").click();");
                    Thread.Sleep(5000);

                    js.ExecuteScript("$(\"#ctl00_cmdSelectMachines\").click();");
                    Thread.Sleep(15000);

                    js.ExecuteScript("$(\"#ctl00_contentPlaceHolderMain_roundedPanelOpPracticesGrid_ImageButtonSummaryGrid\").click();");
                    Thread.Sleep(10000);

                    // We open the HoleSummary file and process every hole
                    HtmlDocument htmlDoc = new HtmlDocument();
                    string uac = "Administrator";
                    string file = @$"C:\Users\{uac}\Downloads\{fileDict[flota]}.xls";
                    htmlDoc.LoadHtml(File.ReadAllText(file));

                    foreach (HtmlNode table in htmlDoc.DocumentNode.SelectNodes("//table"))
                    {
                        List<string> fields = new List<string>();
                        if (table.SelectNodes("thead") == null)
                            continue;
                        foreach (HtmlNode header in table.SelectNodes("thead"))
                        {
                            HtmlNode row = header.SelectNodes("tr")[0];
                            foreach (HtmlNode cell in row.SelectNodes("th"))
                            {
                                if (cell.InnerText != " ")
                                {
                                    var field = cell.InnerText.Replace(" ", "").Replace("(", "").Replace(")", "");
                                    field = RemoveDiacritics(field);
                                    if (field.Contains("Impr"))
                                        field = "TiempoImproductivo";
                                    fields.Add(field);
                                }
                                else
                                {
                                    fields.Add(cell.InnerText);
                                }
                            }
                        }

                        List<string> otherFields = new List<string>();

                        foreach (string fi in fields)
                        {
                            string next = fi;
                            if (next == "Equipment" || next == "Equipo" || next == "Machine")
                                next = "Maquina";
                            else if (next == "MineTime" || next == "Mine Time")
                                next = "TiempoMina";
                            else if (next == "FaultCode" || next == "Fault Code")
                                next = "CodigoFalla";
                            else if (next == "Fault Description" || next == "FaultDescription" || next == "Description")
                                next = "Descripcion";
                            else if (next == "SubSystem")
                                next = "Subsistema";
                            else if (next == "Operator ID")
                                next = "Operador";

                            otherFields.Add(next);
                        }

                        fields = otherFields;

                        foreach (HtmlNode header in table.SelectNodes("tbody"))
                        {
                            if (header.SelectNodes("tr") != null)
                            {
                                foreach (HtmlNode row in header.SelectNodes("tr"))
                                {
                                    int cIndex = 0;

                                    Dictionary<string, object> aData = new Dictionary<string, object>();
                                    foreach (HtmlNode cell in row.SelectNodes("td"))
                                    {
                                        if (fields[cIndex] != "TiempoMina")
                                        {
                                            string val = cell.InnerText.Trim().Replace("&nbsp;", "").Replace("\\r\\n", "");
                                            if (val != "")
                                                aData[fields[cIndex]] = cell.InnerText.Trim().Replace("&nbsp;", "").Replace("\\r\\n", "");
                                            else
                                            {
                                                if (fields[cIndex] == "Gravedad")
                                                    aData[fields[cIndex]] = 0;
                                                else
                                                    aData[fields[cIndex]] = "-";

                                            }
                                        }
                                        else
                                        {
                                            var aDate = CultureInfo.GetCultures(CultureTypes.AllCultures).Select(culture =>
                                            {
                                                DateTime result;
                                                return DateTime.TryParse(
                                                    cell.InnerText.Trim().Replace("&nbsp;", ""),
                                                    culture,
                                                    DateTimeStyles.AssumeLocal,
                                                    out result
                                                ) ? result : default(DateTime?);
                                            })
                                            .Where(d => d != null)
                                            .GroupBy(d => d)
                                            .OrderByDescending(g => g.Count())
                                            .FirstOrDefault()
                                            .Key;
                                            if (aDate.HasValue)
                                            {
                                                aData[fields[cIndex]] = aDate.Value.ToUniversalTime();
                                            }
                                            else
                                            {
                                                aData[fields[cIndex]] = DateTime.MinValue;
                                            }
                                            // aData[fields[cIndex]] = aDate; DateTime.Parse(, null, System.Globalization.DateTimeStyles.AssumeLocal).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
                                        }
                                        cIndex++;
                                    }
                                    aData.Remove("&nbsp;");
                                    aData.Remove("AccionUsuario");
                                    aData.Remove("UserAction");
                                    aData.Remove("Id.Operador");
                                    aData.Remove("Operador");
                                    aData.Remove("Urgency");

                                    failureList.Add(aData);
                                }
                            }
                        }
                    }

                    if (!Directory.Exists(@"C:\PrevailOperating"))
                        Directory.CreateDirectory(@"C:\PrevailOperating");

                    FileInfo f = new FileInfo(file);
                    string movedName = @$"c:\PrevailOperating\processed_{Path.GetFileName(f.FullName)}";
                    if (File.Exists(movedName))
                    {
                        File.Delete(movedName);
                    }
                    File.Move(file, movedName);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"An error ocurred while processing the fleet {flota} information, {e.Message}, {e.StackTrace}");
                }
            }

            List<DataPoint> eventos = new List<DataPoint>();
            try
            {
                foreach (var ev in failureList)
                {
                    var dp = new DataPoint()
                    {
                        ID = 0,
                        office = "La Negra - Chile",
                        company = "Lumina Copper",
                        mine = "Caserones",
                        serialno = ev["Maquina"].ToString(),
                        eventcode = ev["CodigoFalla"].ToString(),
                        eventstarttime = (DateTime)ev["TiempoMina"],
                        eventendtime = DateTime.MinValue,
                        eventdescription = ev["Descripcion"].ToString(),
                        eventidentifier = "",
                        sequenceid = "NA",
                        priority = "",
                        status = "",
                        subsystem = ev["Subsistema"].ToString(),
                        dataloggertimestamp = DateTime.MinValue
                    };
                    eventos.Add(dp);
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            // At the end, we generate a composite package and transmit it to an InputStream.
            var success = TransferPrevailData(eventos);
            if (success)
            {
                Console.WriteLine($"Transfered Failures {failureList.Count}");
            }
            else
            {
                Console.WriteLine($"Failed to Transfer Failures: {failureList.Count}");
            }            
        }

        private static void ProcessFailureInfo(ChromeDriver driver)
        {
            List<string> flotaButton = new List<string>()
            {
                "Drill", "Loader", "Shovel"
            };


            List<Dictionary<string, object>> failureList = new List<Dictionary<string, object>>();

            foreach (string flota in flotaButton)
            {

                try
                {
                    driver.Navigate().GoToUrl($"https://analytics.joyglobal.com/Prevail/{flota}/Maintenance/OutageList.aspx");
                    SwitchToWindow(driver => driver.Title.Contains("Paradas"), driver);
                    Thread.Sleep(5000);

                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("$(\".timeRangeSelectDropDownArrow\")[1].click();");
                    Thread.Sleep(5000);

                    js.ExecuteScript("$(\"#ctl00_radioButtonListDateType_4\").click();");
                    Thread.Sleep(5000);

                    js.ExecuteScript("$(\"#ctl00_cmdSelectMachines\").click();");
                    Thread.Sleep(15000);

                    js.ExecuteScript("$(\"#ctl00_contentPlaceHolderMain_roundedPanelFaultSummaryGrid_imageBtnFaultSummaryGrid\").click();");
                    Thread.Sleep(5000);                    

                    // We open the HoleSummary file and process every hole
                    HtmlDocument htmlDoc = new HtmlDocument();
                    string uac = "Carlos";
                    string file = @$"C:\Users\{uac}\Downloads\FilteredOutageList.xls";
                    htmlDoc.LoadHtml(File.ReadAllText(file));

                    foreach (HtmlNode table in htmlDoc.DocumentNode.SelectNodes("//table"))
                    {
                        List<string> fields = new List<string>();
                        foreach (HtmlNode header in table.SelectNodes("thead"))
                        {
                            HtmlNode row = header.SelectNodes("tr")[0];
                            foreach (HtmlNode cell in row.SelectNodes("th"))
                            {
                                if (cell.InnerText != " ")
                                {
                                    var field = cell.InnerText.Replace(" ", "").Replace("(", "").Replace(")", "");
                                    field = RemoveDiacritics(field);
                                    if (field.Contains("Impr"))
                                        field = "TiempoImproductivo";
                                    fields.Add(field);
                                }
                                else
                                {
                                    fields.Add(cell.InnerText);
                                }
                            }
                        }

                        List<string> otherFields = new List<string>();

                        foreach (string fi in fields)
                        {
                            string next = fi;
                            if (next == "Equipment" || next == "Equipo")
                                next = "Maquina";
                            else if (next == "MineTime")
                                next = "TiempoMina";
                            else if (next == "FaultCode")
                                next = "CodigoFalla";
                            else if (next == "Description")
                                next = "Descripcion";
                            else if (next == "DownTimeMins")
                                next = "TiempoImproductivo";
                            else if (next == "SubSystem")
                                next = "Subsistema";
                            else if (next == "Severity")
                                next = "Gravedad";

                            otherFields.Add(next);
                        }

                        fields = otherFields;

                        foreach (HtmlNode header in table.SelectNodes("tbody"))
                        {
                            if (header.SelectNodes("tr") != null)
                            {
                                foreach (HtmlNode row in header.SelectNodes("tr"))
                                {
                                    int cIndex = 0;

                                    Dictionary<string, object> aData = new Dictionary<string, object>();
                                    foreach (HtmlNode cell in row.SelectNodes("td"))
                                    {
                                        if (fields[cIndex] != "TiempoMina" && fields[cIndex] != "TiempoImproductivo")
                                        {
                                            string val = cell.InnerText.Trim().Replace("&nbsp;", "").Replace("\\r\\n", "");
                                            if (val != "")
                                                aData[fields[cIndex]] = cell.InnerText.Trim().Replace("&nbsp;", "").Replace("\\r\\n", "");
                                            else
                                            {
                                                if (fields[cIndex] == "Gravedad")
                                                    aData[fields[cIndex]] = 0;
                                                else
                                                    aData[fields[cIndex]] = "-";

                                            }
                                        }
                                        else if (fields[cIndex] == "TiempoImproductivo")
                                        {
                                            aData[fields[cIndex]] = double.Parse(cell.InnerText.Trim().Replace("&nbsp;", "").Replace("\\r\\n", ""));
                                        }
                                        else
                                        {
                                            var aDate = CultureInfo.GetCultures(CultureTypes.AllCultures).Select(culture =>
                                            {
                                                DateTime result;
                                                return DateTime.TryParse(
                                                    cell.InnerText.Trim().Replace("&nbsp;", ""),
                                                    culture,
                                                    DateTimeStyles.AssumeLocal,
                                                    out result
                                                ) ? result : default(DateTime?);
                                            })
                                            .Where(d => d != null)
                                            .GroupBy(d => d)
                                            .OrderByDescending(g => g.Count())
                                            .FirstOrDefault()
                                            .Key;
                                            if (aDate.HasValue)
                                            {
                                                aData[fields[cIndex]] = aDate.Value.ToUniversalTime();
                                            }
                                            else
                                            {
                                                aData[fields[cIndex]] = DateTime.MinValue;
                                            }
                                            // aData[fields[cIndex]] = aDate; DateTime.Parse(, null, System.Globalization.DateTimeStyles.AssumeLocal).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
                                        }
                                        cIndex++;
                                    }
                                    aData.Remove("&nbsp;");
                                    aData.Remove("AccionUsuario");
                                    aData.Remove("UserAction");
                                    aData.Remove("Id.Operador");
                                    aData.Remove("OperatorId");
                                    aData.Remove("Urgency");

                                    failureList.Add(aData);
                                }
                            }
                        }
                    }

                    if (!Directory.Exists(@"C:\PrevailFailure"))
                        Directory.CreateDirectory(@"C:\PrevailFailure");

                    FileInfo f = new FileInfo(file);
                    string movedName = @$"c:\PrevailFailure\processed_{Path.GetFileName(f.FullName)}";
                    if (File.Exists(movedName))
                    {
                        File.Delete(movedName);
                    }
                    File.Move(file, movedName);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"An error ocurred while processing the fleet {flota} information, {e.Message}, {e.StackTrace}");
                }
            }

            // At the end, we generate a composite package and transmit it to an InputStream.
            var success = TransferFailureData(failureList, "");
            if (success)
            {
                Console.WriteLine($"Transfered Failures {failureList.Count}");
            }
            else
            {
                Console.WriteLine($"Failed to Transfer Failures: {failureList.Count}");
            }
            driver.Navigate().GoToUrl("https://analytics.joyglobal.com/Prevail/Drill/Production/HoleSummary.aspx");
        }

        public static bool IterateDRDashboard(ChromeDriver driver, string equipment)
        {
            ICollection<GrafanaQuery> queries = new List<GrafanaQuery>()
            {
                new GrafanaQuery()
                {
                    Alias = "DRWaterInjDriveSpeedRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_WaterSpdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRWaterInjDriveSpeedEstimate",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_WaterActSpd",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRWaterInjMotorCurrent",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_WaterMtrCur",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRWaterInjDCVoltage",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_WaterDCVoltage",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRWaterInjHeatSinkTemp",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_WaterHeatSink_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempRotary2DriveFahrenheit",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_Rot2DriveTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempRotary2MtrFieldFahren",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_Rot2FieldTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempRotary2MtrInterpoleFa",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_Rot2InterTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempHoistDriveFahrenheit",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_HstDriveTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempHoisMtrFieldFahrenheit",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_HstFieldTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempHoisMtrInterpoloFahrenheit",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_HstInterTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DRTempHoistBlowerMtrRunning",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_HstBlwrMtrRun",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_MainAirPSI",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_MainAirPSI",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_LowerGreasePS",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_LowerGreasePS",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_UpperGreasePS",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_UpperGreasePS",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                 new GrafanaQuery()
                {
                    Alias = "DR_HydrTankTemp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_HydrTankTemp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_T_TankOilTemp_C",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_T_TankOilTemp_C",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_T_TankOilTemp_F",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_T_TankOilTemp_F",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_CoolerMtrSclSpd",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_CoolerMtrSclSpd",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_CoolerInletPSI",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_CoolerInletPSI",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_CoolerOutletPSI",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_CoolerOutletPSI",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_PrplFwdSpdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_PrplFwdSpdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_PrplRevSpdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_PrplRevSpdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_PrplLeftSpdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_PrplLeftSpdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_PrplRightSpdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_PrplRightSpdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_LftPrpelFwdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_LftPrpelFwdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_LftPrpelRevRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_LftPrpelRevRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_RghtPrpelFwdRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_RghtPrpelFwdRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "DR_RghtPrpelRevRef",
                    Query = new {
                        aggregator = "sum",
                        metric = "DR_RghtPrpelRevRef",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "1m-avg"
                            }
                        }
                    }
                }



            };


            List<GrafanaMeasurement> measurements = new List<GrafanaMeasurement>();
            DateTime start = DateTime.UtcNow.AddHours(-1.1);
            foreach (GrafanaQuery query in queries)
            {
                measurements.AddRange(GetTrendData(driver, query, start, @$"C:\PrevailIndex\Grafana\Drills"));
            }
            
            if (measurements.Count > 0)
            {
                bool insertResult = TransferGrafanaData(measurements, "f7c9aa36c33bd21551e40767460c7d1d2222d5bf");
                if (insertResult)
                {
                    Console.WriteLine($"Sent {measurements.Count}");
                }
            }
            else
            {
                Console.WriteLine("No new data available on Grafana");
            }
            return true;

        }

        public static bool IterateLDDashboard(ChromeDriver driver, string equipment)
        {
            ICollection<GrafanaQuery> queries = new List<GrafanaQuery>()
            {
                new GrafanaQuery()
                {
                    Alias = "LDSumBucketLoads",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Bucket_Load",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDMaxBucketLoad",
                    Query = new {
                        aggregator = "max",
                        metric = "LD_Bucket_Load",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDAvgBucketLoad",
                    Query = new {
                        aggregator = "avg",
                        downsample = "5m-avg",
                        metric = "LD_Bucket_Load",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDCountBucketLoad",
                    Query = new {
                        aggregator = "count",
                        metric = "LD_Bucket_Load",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDHydReservoirOilMaxTemp",
                    Query = new {
                        metric = "LD_Hyd_Reservoir_Oil",
                        aggregator = "max",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDHydReservoirOilTemp",
                    Query = new {
                        metric = "LD_Hyd_Reservoir_Oil",
                        aggregator = "avg",
                        downsample = "1m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDGearboxOilMaxTemp",
                    Query = new {
                        metric = "LD_HPD_Gearbox_Oil",
                        aggregator = "max",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDGearboxOilTemp",
                    Query = new {
                        metric = "LD_HPD_Gearbox_Oil",
                        aggregator = "avg",
                        downsample = "1m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDEngineOilTemp",
                    Query = new {
                        metric = "LD_Engine_Oil_Temp",
                        aggregator = "avg",
                        downsample = "1m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    },
                },
                new GrafanaQuery()
                {
                    Alias = "LDSystemAirPress",
                    Query = new {
                        metric = "LD_System_Air",
                        aggregator = "avg",
                        downsample = "2m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDHydReservioirPress",
                    Query = new {
                        metric = "LD_Hyd_Reservoir",
                        aggregator = "avg",
                        downsample = "2m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    },
                },
                new GrafanaQuery()
                {
                    Alias = "LDAutolubeGreasePress",
                    Query = new {
                        metric = "LD_Autolube_Grease",
                        aggregator = "avg",
                        downsample = "2m-avg",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    },
                },
                new GrafanaQuery()
                {
                    Alias = "LDHourmeter",
                    Query = new {
                        metric = "LD_Machine",
                        aggregator = "max",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LDArticulatedFaceEngagementSTS",
                    Query = new {
                        metric = "LD_Articulated_Face_Engagement_STS",
                        aggregator = "max",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Engine_Oil_Temp",
                    Query = new {
                        metric = "LD_Engine_Oil_Temp",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Engine_Coolant_Temp",
                    Query = new {
                        metric = "LD_Engine_Coolant_Temp",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_GenMst_Greatest_IGBT",
                    Query = new {
                        metric = "LD_GenMst_Greatest_IGBT",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_LF_Mst_Greatest_IGBT",
                    Query = new {
                        metric = "LD_LF_Mst_Greatest_IGBT",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_RF_Mst_Geatest_IGBT",
                    Query = new {
                        metric = "LD_RF_Mst_Geatest_IGBT",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_RR_Mst_Greatest_IGBT",
                    Query = new {
                        metric = "LD_RR_Mst_Greatest_IGBT",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_LR_Mst_Greatest_IGBT",
                    Query = new {
                        metric = "LD_LR_Mst_Greatest_IGBT",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Blower",
                    Query = new {
                        metric = "LD_Blower",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Blower_Pump",
                    Query = new {
                        metric = "LD_Blower_Pump",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Radiator_Fan",
                    Query = new {
                        metric = "LD_Radiator_Fan",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Radiator_Fan_Pump",
                    Query = new {
                        metric = "LD_Radiator_Fan_Pump",
                        aggregator = "sum",
                        filters = new List<dynamic>() {
                            new {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or",
                                downsample = "5m-avg"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Battery",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Battery",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Radiator_Fan",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Radiator_Fan",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Radiator_Fan_Pump",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Radiator_Fan_Pump",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Engine_Fuel_Temp",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Engine_Fuel_Temp",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                },
                new GrafanaQuery()
                {
                    Alias = "LD_Engine_Coolant_Temp",
                    Query = new {
                        aggregator = "sum",
                        metric = "LD_Engine_Coolant_Temp",
                        filters = new List<dynamic>()
                        {
                            new
                            {
                                filter = equipment,
                                groupBy = false,
                                tagk = "SerialNo",
                                type = "literal_or"
                            }
                        }
                    }
                }
            };

            List<GrafanaMeasurement> measurements = new List<GrafanaMeasurement>();
            DateTime start = DateTime.UtcNow.AddHours(-3.5);
            DateTime startAFE = DateTime.UtcNow.AddDays(-4);
            foreach (GrafanaQuery query in queries)
            {
                if (query.Alias == "LDArticulatedFaceEngagementSTS")
                {
                    var meas = GetTrendData(driver, query, startAFE, @$"C:\PrevailIndex\Grafana\Loaders");
                    measurements.AddRange(meas);
                }
                else
                {
                    measurements.AddRange(GetTrendData(driver, query, start, @$"C:\PrevailIndex\Grafana\Loaders"));
                }
            }

            if (measurements.Count > 0)
            {
                bool insertResult = TransferGrafanaData(measurements, "490851463b85c2046976658f6f0e051690a0a30f");
                if (insertResult)
                {

                    Console.WriteLine($"Sent {measurements.Count}");
                }
            }
            else
            {
                Console.WriteLine("No new data available on Grafana");
            }

            string ruta = @$"C:\PrevailIndex\Grafana\Loaders\spd_hr_{DateTime.Now.ToString("yyyy-MM-dd-HH")}-{equipment}";

            if(!File.Exists(ruta))
            {
                var newIndex = File.Create(ruta);
                newIndex.Close();

                ICollection<GrafanaQuery> queryVelocidades = new List<GrafanaQuery>()
                {
                    // Integración para Analisis de Velocidad de Carguío new GrafanaQuery()
                    new GrafanaQuery()
                    {
                        Alias = "LDArticulatedFaceEngagementSTS",
                        Query = new {
                            metric = "LD_Articulated_Face_Engagement_STS",
                            aggregator = "max",
                            filters = new List<dynamic>() {
                                new {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Bucket_Position",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Bucket_Position",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Lift_Arm_Position",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Lift_Arm_Position",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Steer_Position",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Steer_Position",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Vehicle",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Vehicle",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Spd_Control_Pedal",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Spd_Control_Pedal",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_JstickL_Direction_Switch",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_JstickL_Direction_Switch",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    },
                    new GrafanaQuery()
                    {
                        Alias = "LD_Bucket_Load",
                        Query = new {
                            aggregator = "sum",
                            downsample = "10ms-avg",
                            metric = "LD_Bucket_Load",
                            filters = new List<dynamic>()
                            {
                                new
                                {
                                    filter = equipment,
                                    groupBy = false,
                                    tagk = "SerialNo",
                                    type = "literal_or"
                                }
                            }
                        }
                    }
                };

                List<GrafanaMeasurement> velocidadesMeasurements = new List<GrafanaMeasurement>();
                DateTime startSpeed = DateTime.UtcNow.AddHours(-12);
                foreach (GrafanaQuery query in queryVelocidades)
                {
                    DateTime ss = startSpeed.AddMinutes(0);
                    DateTime end = DateTime.UtcNow;
                    while (ss <= end)
                    {
                        velocidadesMeasurements.AddRange(GetTrendData(driver, query, ss, ss.AddHours(1), @$"C:\PrevailIndex\Grafana\Loaders"));
                        ss = ss.AddHours(1);
                    }
                }

                if (velocidadesMeasurements.Count > 0)
                {
                    bool insertResult = TransferGrafanaData(velocidadesMeasurements, "42e93162e5a2e70b1890ef5e96b5a4cc52f9de70");
                    if (insertResult)
                    {

                        Console.WriteLine($"Sent {velocidadesMeasurements.Count} for speed analysis");
                    }
                }
                else
                {
                    Console.WriteLine("No new data available on Grafana");
                }
            }

            
            return true;
        }

        private static List<GrafanaMeasurement> GetTrendData(ChromeDriver driver, GrafanaQuery thequery, DateTime start, DateTime end, string path)
        {
            long unixTimestamp = (long)(start.Subtract(new DateTime(1970, 1, 1))).TotalSeconds * 1000;
            long unixTimestampEnd = (long)(end.Subtract(new DateTime(1970, 1, 1))).TotalSeconds * 1000;
            dynamic query = new
            {
                globalAnnotations = true,
                msResolution = true,
                start = unixTimestamp,
                end = unixTimestampEnd,
                queries = new List<dynamic>()
                {
                    thequery.Query
                }
            };

            string queryString = JsonConvert.SerializeObject(query);

            string theFx = "async function go() { " +
                           " var result = []; " +
                                    "var code = 0;" +
                                    "var message = 'OK';" +
                                    "var retorno = {};" +
                            "try{" +
                            "await $.ajax({url: 'https://defaulttrends.joyglobal.com/api/datasources/proxy/4/api/query'," +
                            " headers:{'X-Grafana-Org-Id': 2,'X-Grafana-User-ID': 1604,'Content-Type': 'application/json;charset=UTF-8'," +
                                    "'Accept': 'application/json, text/plain, /','Accept-Language': 'en-US,en;q=0.9'}," +
                            " type: 'POST',datatype: 'json',data: JSON.stringify(" + queryString + ")," +
                            " success: function(data, status, xhr){result = data;code = 200;}," +
                            " error: function(xhr, status, error){ message = xhr.responseText;code = xhr.status;} });" +
                            "}catch (err){}" +
                             " retorno = { 'code': code, 'message': message, 'data': result};" +
                              " return retorno;" +
                             "}" +
                             "return await go();";
            List<GrafanaMeasurement> measurements = new List<GrafanaMeasurement>();
            Stopwatch sw = new Stopwatch();
            object result;
            GrafanaAjaxResult returnAjax = new GrafanaAjaxResult();
            ControlPrevail cp = new ControlPrevail();
            try
            {
                sw.Start();
                result = driver.ExecuteScript(theFx);



                returnAjax = JsonConvert.DeserializeObject<GrafanaAjaxResult>(JsonConvert.SerializeObject(result));

                if (returnAjax.Message.Contains("Unauthorized"))
                {
                    throw new Exception("Grafana closed the connection");
                }

                foreach (var o in returnAjax.Data)
                {
                    IDictionary<string, object> vals = o.ToObject<Dictionary<string, object>>();
                    var kpiValues = (vals["dps"] as Newtonsoft.Json.Linq.JObject).ToObject<Dictionary<string, object>>();
                    string metric = vals["metric"] as string;
                    var tags = (vals["tags"] as JObject).ToObject<Dictionary<string, object>>();
                    string source = (tags["DataSource"].GetType().Name == "String") ? tags["DataSource"].ToString() : ((JObject)tags["DataSource"]).ToObject<string>();
                    var dataQuality = tags["DataQuality"] as string;
                    var serial = tags["SerialNo"] as string;
                    var dataType = tags["DataType"] as string;
                    foreach (var timeKey in kpiValues.Keys)
                    {
                        double tk = double.Parse(timeKey);
                        DateTime readOn = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Local);
                        readOn = readOn.AddMilliseconds(tk);

                        while (readOn > DateTime.Now)
                            readOn = readOn.AddHours(-1);

                        GrafanaMeasurement gm = new GrafanaMeasurement()
                        {
                            SerialNo = serial,
                            ReadOn = readOn.ToUniversalTime(),
                            Metric = metric,
                            DataQuality = dataQuality,
                            DataSource = source,
                            DataType = dataType,
                            Value = kpiValues[timeKey],
                            Alias = thequery.Alias
                        };
                        measurements.Add(gm);
                    }
                }

                long cycles = measurements.Count(a => Convert.ToInt32(a.Value) == 1);
                if (measurements.Count > 0)
                    Console.WriteLine(measurements.Max(x => x.ReadOn.ToLocalTime()).ToString("s"));
                sw.Stop();

            }
            catch (Exception e)
            {
                //DoSendEmail($"Ocurrió un error al realizar la consulta {thequery.Alias} <br> Exception : {e.StackTrace} <br> Exception MSG: {e.Message}", "Grafana - Default Trends Home");
                cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, 0, sw.ElapsedMilliseconds / 1000, $"Cod error: Unserializable Response or Processing Error");
                ControlPrevailList.Add(cp);
            }

            if (returnAjax.Message == null)
            {
                cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, measurements.Count, (float)sw.Elapsed.TotalSeconds, "Timeout al llamar a la consulta");
            }
            else
            {
                if (returnAjax.Message.Length > 2)
                {
                    cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, measurements.Count, (float)sw.Elapsed.TotalSeconds, $"Cod error: {returnAjax.Code} <br> MSG: {System.Uri.EscapeDataString(returnAjax.Message)}");
                }
                else
                {
                    cp = ControlPrevailBuilder("Grafana", true, thequery.Alias, measurements.Count, (float)sw.Elapsed.TotalSeconds, $"Ok");
                }
            }

            ControlPrevailList.Add(cp);
            return measurements;
        }

        private static List<GrafanaMeasurement> GetTrendData(ChromeDriver driver, GrafanaQuery thequery, DateTime start, string path)
        {
            long unixTimestamp = (long)(start.Subtract(new DateTime(1970, 1, 1))).TotalSeconds * 1000;
            //long unixTimestampEnd = (long)(end.Subtract(new DateTime(1970, 1, 1))).TotalSeconds * 1000;
            dynamic query = new
            {
                globalAnnotations = true,
                msResolution = true,
                start = unixTimestamp,
                //end = unixTimestampEnd,
                queries = new List<dynamic>()
                {
                    thequery.Query
                }
            };

            string queryString = JsonConvert.SerializeObject(query);

            string theFx = "async function go() { " +
                           " var result = []; " +
                                    "var code = 0;" +
                                    "var message = 'OK';" +
                                    "var retorno = {};" +
                            "try{" +
                            "await $.ajax({url: 'https://defaulttrends.joyglobal.com/api/datasources/proxy/4/api/query'," +
                            " headers:{'X-Grafana-Org-Id': 2,'X-Grafana-User-ID': 1604,'Content-Type': 'application/json;charset=UTF-8'," +
                                    "'Accept': 'application/json, text/plain, /','Accept-Language': 'en-US,en;q=0.9'}," +
                            " type: 'POST',datatype: 'json',data: JSON.stringify(" + queryString + ")," +
                            " success: function(data, status, xhr){result = data;code = 200;}," +
                            " error: function(xhr, status, error){ message = xhr.responseText;code = xhr.status;} });" +
                            "}catch (err){}" +
                             " retorno = { 'code': code, 'message': message, 'data': result};" +
                              " return retorno;" +
                             "}" +
                             "return await go();";
            List<GrafanaMeasurement> measurements = new List<GrafanaMeasurement>();
            Stopwatch sw = new Stopwatch();
            object result;
            GrafanaAjaxResult returnAjax = new GrafanaAjaxResult();
            ControlPrevail cp = new ControlPrevail();
            try
            {
                sw.Start();
                result = driver.ExecuteScript(theFx);
               
                returnAjax = JsonConvert.DeserializeObject<GrafanaAjaxResult>(JsonConvert.SerializeObject(result));

                if (returnAjax.Message.Contains("Unauthorized"))
                {
                    throw new Exception("Grafana closed the connection");
                }

                foreach (var o in returnAjax.Data)
                {
                    IDictionary<string, object> vals = o.ToObject<Dictionary<string, object>>();
                    var kpiValues = (vals["dps"] as Newtonsoft.Json.Linq.JObject).ToObject<Dictionary<string, object>>(); // vals["dps"] as IDictionary<string, object>;
                    string metric = vals["metric"] as string;
                    var tags = (vals["tags"] as Newtonsoft.Json.Linq.JObject).ToObject<Dictionary<string, string>>();
                    var source = tags["DataSource"] as string;
                    var dataQuality = tags["DataQuality"] as string;
                    var serial = tags["SerialNo"] as string;
                    var dataType = tags["DataType"] as string;
                    foreach (var timeKey in kpiValues.Keys)
                    {
                        double tk = double.Parse(timeKey);
                        DateTime readOn = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Local);
                        readOn = readOn.AddMilliseconds(tk);

                        while (readOn > DateTime.Now)
                            readOn = readOn.AddHours(-1);
                        
                        GrafanaMeasurement gm = new GrafanaMeasurement()
                        {
                            SerialNo = serial,
                            ReadOn = readOn.ToUniversalTime(),
                            Metric = metric,
                            DataQuality = dataQuality,
                            DataSource = source,
                            DataType = dataType,
                            Value = kpiValues[timeKey],
                            Alias = thequery.Alias
                        };
                        measurements.Add(gm);
                    }
                }

                long cycles = measurements.Count(a => Convert.ToInt32(a.Value) == 1);
                if(measurements.Count > 0)
                    Console.WriteLine(measurements.Max(x=> x.ReadOn.ToLocalTime()).ToString("s"));
                sw.Stop();

            }
            catch (Exception e)
            {
                //DoSendEmail($"Ocurrió un error al realizar la consulta {thequery.Alias} <br> Exception : {e.StackTrace} <br> Exception MSG: {e.Message}", "Grafana - Default Trends Home");
                cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, 0, sw.ElapsedMilliseconds/1000, $"Cod error: Unserializable Response or Processing Error");
                ControlPrevailList.Add(cp);
            }

            if(returnAjax.Message == null)
            {
                cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, measurements.Count, (float)sw.Elapsed.TotalSeconds, "Timeout al llamar a la consulta");
            }
            else
            {
                if (returnAjax.Message.Length > 2)
                {
                    cp = ControlPrevailBuilder("Grafana", false, thequery.Alias, measurements.Count, (float) sw.Elapsed.TotalSeconds, $"Cod error: {returnAjax.Code} <br> MSG: {System.Uri.EscapeDataString(returnAjax.Message)}");
                }
                else
                {
                    cp = ControlPrevailBuilder("Grafana", true, thequery.Alias, measurements.Count, (float)sw.Elapsed.TotalSeconds , $"Ok");
                }
            }
             
            ControlPrevailList.Add(cp);
            return measurements;
        }

        private static bool ProcessEquipmentData(ChromeDriver driver, string dataOption)
        {
            try
            {
                IWebElement btnDropdown = driver.FindElement(By.ClassName("k-input"));
                btnDropdown.Click();

                Thread.Sleep(2000);

                var opts = driver.FindElements(By.ClassName("k-item"));
                foreach (var option in opts)
                {
                    if (option.Text == dataOption) //"ae_ld_view")
                    {
                        option.Click();
                    }
                }

                IWebElement btn = driver.FindElement(By.Id("AEExportCSV"));
                btn.Click();

                Thread.Sleep(15000);

                ProcessCsvData(dataOption);

                return false;
            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al procesar AE {dataOption} <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "AE Live");
            }
            return false;
        }


        private static bool ProcessEquipmentJsonData(ChromeDriver driver, string dataOption)
        {
            CheckAlert(driver);
            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("$('.k-dropdown')[0].click();");
                IWebElement btnDropdown = driver.FindElement(By.ClassName("k-dropdown"));
                Thread.Sleep(5000);
                js.ExecuteScript("var items = $('.k-item'); $.each(items, function(i, v) {if($(v).text() == \"" + dataOption + "\") $(v).click(); });");
               
                Thread.Sleep(5000);

                string theFx = "async function go(){ " +
                                "var value = []; " +
                                "var datacontext = ''; " +
                                "var datacount = 0; " +
                                "var code = 0; " +
                                "var message = 'OK'; " +
                                "var retorno = { }; " +
                                "var retornoAE = { }; " +
                                "try {await $.ajax({ url: 'https://analytics.joyglobal.com/odata/alarmevents?$inlinecount=allpages&take=10000&skip=0&page=1&pageSize=100000', " +
                                "headers:{ " +
                                "'Content-Type': 'application/json;charset=UTF-8', " +
                                "'Accept': '*/*', " +
                                "'Accept-Language': 'en-US,en;q=0.9'}, " +
                                "type: 'GET', datatype: 'json', " +
                                "success: function(data, status, xhr) {" +
                                "retornoAE = data; code = 200; }, " +
                                "error: function(xhr, status, error) { " +
                                "message = xhr.responseText; code = xhr.status; " +
                                "retornoAE = {'value' : [], 'datacontext' : 'error','datacount' : 0 }; }}); " +
                                "} catch (err) {}" +
                                "retorno = {'code': code,'message': message,'value': retornoAE['value'],'datacontext' : retornoAE['@odata.context'],'datacount' : retornoAE['@odata.count']" +
                                ", 'dataoption' : '" + dataOption + "' " +
                                "}; return retorno; } return await go();";
                Object result = driver.ExecuteScript(theFx);

                Thread.Sleep(5000);
                using (StreamWriter file = File.CreateText(@$"C:\PrevailJson\{dataOption}_{DateTime.UtcNow.ToString("yyyy_MM_dd")}.json"))
                {
                    Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer();
                    //serialize object directly into file stream
                    serializer.Serialize(file, result);
                    file.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al procesar AE {dataOption} <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "AE Live");
            }
            return false;
        }

        private static bool ProcessJsonData()
        {
            MainBody jsonBody = new MainBody();
            ControlPrevail cp = new ControlPrevail();
            int foundData = 0;
            int sentData = 0;
            Stopwatch sw = new Stopwatch();
            try
            {
                sw.Start();
                ICollection<string> files = new List<string>(Directory.EnumerateFiles(@"C:\PrevailJson"));
                
                if (files.Count > 0)
                {
                    foreach (string f in files)
                    {
                        using (StreamReader file = File.OpenText(f))
                        {
                            Newtonsoft.Json.JsonSerializer serializer = new Newtonsoft.Json.JsonSerializer()
                            {
                                DateTimeZoneHandling = DateTimeZoneHandling.Local
                            };

                            jsonBody = (MainBody)serializer.Deserialize(file, typeof(MainBody));

                            List<DataPoint> forTransmission = new List<DataPoint>();
                            if (jsonBody.value != null)
                            {
                                foreach (var dp in jsonBody.value)
                                {
                                    /*
                                    if(dp.serialno.Contains("ES"))
                                        dp.eventstarttime = dp.eventstarttime.AddHours(3);
                                    else
                                        dp.eventstarttime = dp.eventstarttime.AddHours(4);
                                    */

                                    /*
                                     * BD320191 -4
                                        BD320177 -3 (Hora antigua, muestra 1 hora adelante en Grafana)
                                        BD320201 (Sin datos)

                                        LD-2350-2209 -4
                                        LD-2350-2223 -4

                                        ES41269 -3 (Hora antigua, muestra 1 hora adelante en Grafana)
                                     */
                                    
                                    if (dp.serialno == "ES41269" || dp.serialno == "BD320177")
                                    {
                                        dp.eventstarttime = dp.eventstarttime.AddHours(3);
                                    }
                                    else
                                    {
                                        dp.eventstarttime = dp.eventstarttime.AddHours(4);
                                    }

                                    dp.eventstarttime = dp.eventstarttime.ToUniversalTime();
                                    forTransmission.Add(dp);
                                }
                            }

                            
                            

                           forTransmission = forTransmission.GroupBy(x => new { x.serialno, x.eventcode, x.eventstarttime })
                                      .Select(g => g.OrderByDescending(o => o.eventstarttime).First())
                                      .ToList();

                            if (forTransmission.Count > 0)
                            {
                                bool success = TransferPrevailData(forTransmission);
                                if (success)
                                {                                    
                                    file.Close();
                                    string movedName = @$"c:\PrevailProcessedJson\processed_{Path.GetFileName(f)}";
                                    if (File.Exists(movedName))
                                    {
                                        File.Delete(movedName);
                                    }
                                    File.Move(f, movedName);
                                }
                            }
                            else
                            {
                                Console.WriteLine($"No new data available in Prevail Json");

                                file.Close();
                                string movedName = @$"c:\PrevailProcessedJson\processed_{Path.GetFileName(f)}";
                                if (File.Exists(movedName))
                                {
                                    File.Delete(movedName);
                                }
                                File.Move(f, movedName);
                            }
                        }
                    }
                }
                sw.Stop();
                if (jsonBody.message == null)
                {
                    cp = ControlPrevailBuilder("AE", false, jsonBody.dataoption, 0, (float)sw.Elapsed.TotalSeconds, $"Timeout al llamar a la consulta");
                }
                else
                {
                    if (jsonBody.message.Length > 2)
                    {
                        cp = ControlPrevailBuilder("AE", false, jsonBody.dataoption, sentData, (float)sw.Elapsed.TotalSeconds, $"Cod error: {jsonBody.code} <br> Exception MSG: {System.Uri.EscapeDataString(jsonBody.message)}");
                    }
                    else
                    {
                        cp = ControlPrevailBuilder("AE", true, jsonBody.dataoption, sentData, (float)sw.Elapsed.TotalSeconds , $"Ok");
                    }
                }
                Console.WriteLine($"Sent from AE {sentData} / {foundData}");
                ControlPrevailList.Add(cp);
            }
            catch (Exception ex)
            {                
                cp = ControlPrevailBuilder("AE", false, jsonBody.dataoption, 0, (float)sw.Elapsed.TotalSeconds, $"Error: {ex.Message} <br> StackTrace: {ex.StackTrace}");
                ControlPrevailList.Add(cp);
                return false; 
            }
            
            return true;
        }

        private static bool ProcessCsvData(string dataOption)
        {
            ICollection<string> files = new List<string>(Directory.EnumerateFiles(@"C:\Prevail"));
            int foundData = 0;
            int sentData = 0;
            if (files.Count > 0)
            {
                foreach (string file in files)
                {
                    ICollection<DataPoint> data = ParsePoints(file);
                    List<DataPoint> forTransmission = new List<DataPoint>();


                    foreach (var dp in data)
                    {

                        string serial = dp.serialno;
                        DateTime insertTime = dp.inserttime;
                        string code = dp.eventcode.Replace("<", "").Replace(">", "");
                        string index = $"{serial}_{code}_{insertTime.ToString("hhmmss")}";
                        foundData++;
                        if (!Directory.Exists(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\"))
                            Directory.CreateDirectory(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}");

                        var r = File.Exists(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\{index}.csv");
                        // No existe, transmitir y almacenar indice
                        if (r == false)
                        {
                            forTransmission.Add(dp);
                        }
                    }

                    if (forTransmission.Count > 0)
                    {
                        bool success = TransferPrevailData(forTransmission);
                        if (success)
                        {
                            foreach (DataPoint dp in forTransmission)
                            {
                                string serial = dp.serialno;
                                DateTime insertTime = dp.inserttime;
                                string code = dp.eventcode.Replace("<", "").Replace(">", "");
                                string index = $"{serial}_{code}_{insertTime.ToString("hhmmss")}";
                                var r = File.Create(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\{index}.csv");
                                r.Close();
                                sentData++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"No new data available in Prevail for {dataOption}");
                    }


                    string movedName = @$"c:\PrevailProcessed\processed_{dataOption}_{DateTime.UtcNow.ToString("yyyyMMddhhmmss")}.csv";
                    if (File.Exists(movedName))
                    {
                        File.Delete(movedName);
                    }
                    File.Move(file, movedName);

                }

                Console.WriteLine($"Sent {sentData} / {foundData} for {dataOption}");
                return true;
            }
            return true;
        }

        private static bool BackProcessCsvData(string dataOption)
        {
            ICollection<string> files = new List<string>(Directory.EnumerateFiles(@"C:\Prevail"));
            int foundData = 0;
            int sentData = 0;
            if (files.Count > 0)
            {
                foreach (string file in files)
                {
                    if (true)
                    {
                        DateTime dTime = new DateTime(2021, 04, 03);
                        DateTime eTime = new DateTime(2021, 04, 10);
                        ICollection<DataPoint> data = ParsePoints(file);

                        foreach (DataPoint dp in data)
                        {
                            if (dp.eventstarttime >= dTime)
                            {
                                if (dp.serialno == "ES41269" || dp.serialno == "BD320177")
                                {
                                    dp.eventstarttime = dp.eventstarttime.AddHours(3);
                                }
                                else
                                {
                                    dp.eventstarttime = dp.eventstarttime.AddHours(4);
                                }
                            }
                            else
                            {
                                dp.eventstarttime = dp.eventstarttime.AddHours(3);
                            }
                        }

                        List<DataPoint> forTransmission = new List<DataPoint>();


                        foreach (var dp in data)
                        {

                            string serial = dp.serialno;
                            DateTime insertTime = dp.inserttime;
                            string code = dp.eventcode.Replace("<", "").Replace(">", "");
                            string index = $"{serial}_{code}_{insertTime.ToString("hhmmss")}";
                            foundData++;
                            if (!Directory.Exists(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\"))
                                Directory.CreateDirectory(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}");

                            var r = false; // File.Exists(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\{index}.csv");
                            // No existe, transmitir y almacenar indice
                            if (r == false)
                            {
                                forTransmission.Add(dp);
                            }
                        }

                        if (forTransmission.Count > 0)
                        {
                            bool success = TransferPrevailData(forTransmission);
                            //if (success)
                            //{
                            //    foreach (DataPoint dp in forTransmission)
                            //    {
                            //        string serial = dp.serialno;
                            //        DateTime insertTime = dp.inserttime;
                            //        string code = dp.eventcode.Replace("<", "").Replace(">", "");
                            //        string index = $"{serial}_{code}_{insertTime.ToString("hhmmss")}";
                            //        var r = File.Create(@$"C:\PrevailIndex\AE\{insertTime.ToString("yyyyMMddHH")}\{index}.csv");
                            //        r.Close();
                            //        sentData++;
                            //    }
                            //}
                        }
                        else
                        {
                            Console.WriteLine($"No new data available in Prevail for {dataOption}");
                        }


                        string movedName = @$"c:\PrevailProcessed\processed_{dataOption}_{DateTime.UtcNow.ToString("yyyyMMddhhmmss")}.csv";
                        if (File.Exists(movedName))
                        {
                            File.Delete(movedName);
                        }
                        File.Move(file, movedName);
                    }

                }

                Console.WriteLine($"Sent {sentData} / {foundData} for {dataOption}");
                return true;
            }
            return true;
        }

        private static ICollection<DataPoint> ParsePoints(string file)
        {
            StreamReader sr = new StreamReader(file);
            ICollection<DataPoint> data = new List<DataPoint>();
            while (!sr.EndOfStream)
            {
                string linea = sr.ReadLine();
                if (!linea.Contains("office,company,mine"))
                {
                    string[] comps = linea.Split(",");
                    /* Structure
                     * office,
                     * company,
                     * mine,
                     * serialno,
                     * eventstarttime,
                     * eventcode,
                     * EventIdentifier,
                     * eventdescription,
                     * subsystem,
                     * Component,
                     * value,
                     * fromvalue,
                     * EventType,
                     * status,
                     * SequenceID,
                     * inserttime,
                     * OperatorID,
                     * graphfilepath
                     */

                                    DataPoint dp = new DataPoint()
                    {
                        office = comps[0],
                        company = comps[1],
                        mine = comps[2],
                        serialno = comps[3],
                        eventstarttime = DateTime.Parse(comps[4]),
                        eventcode = comps[5],
                        eventidentifier = comps[6],
                        eventdescription = comps[7],
                        subsystem = comps[8],
                        component = comps[9],
                        value = comps[10],
                        fromvalue = comps[11],
                        eventtype = comps[12],
                        status = comps[13],
                        sequenceid = comps[14],
                        inserttime = DateTime.Parse(comps[15]),
                        OperatorID = comps[16]
                    };
                    data.Add(dp);

                }
            }
            sr.Close();
            return data;
        }

        private static bool TransferGrafanaData(List<GrafanaMeasurement> gMetrics, string key)
        {
            int currentIndex = 0;
            bool success = true;
            try
            {
                while (currentIndex < gMetrics.Count)
                {
                    int count = 10000;
                    if (currentIndex + 1000 >= gMetrics.Count)
                    {
                        count = gMetrics.Count - currentIndex;
                    }
                    Console.WriteLine($"{currentIndex} to {currentIndex + count} of {gMetrics.Count}");
                    ICollection<GrafanaMeasurement> sub = gMetrics.GetRange(currentIndex, count);
                    currentIndex = currentIndex + count;
                    Dictionary<string, string> allData = new Dictionary<string, string>();
                    allData.Add("IKEY", key);
                    allData.Add("Data", JsonConvert.SerializeObject(sub));
                    allData.Add("Mode", "Bulk");

                    /* Here we initialize the RestClient. Remember to replace your API Key. */
                    RestClient restClient = new RestClient("https://listen.a2g.io/v1/production/microstream");
                    RestRequest restRequest = new RestRequest(Method.POST);
                    restRequest.AddHeader("x-api-key", "");
                    restRequest.AddJsonBody(allData);
                    restRequest.Timeout = 120000;

                    IRestResponse response = restClient.Execute(restRequest);
                    if (response.IsSuccessful == false)
                    {
                        success = false;
                        Console.WriteLine(response.Content);
                    }
                }

            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al enviar la data <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "Grafana - Default Trends Home");
            }
            return success;
        }

        private static bool TransferPrevailAlert(List<ControlPrevail> cp)
        {
            bool success = true;
            try
            {
                    Dictionary<string, string> allData = new Dictionary<string, string>();
                    allData.Add("IKEY", "");
                    allData.Add("Data", JsonConvert.SerializeObject(cp));
                    allData.Add("Mode", "Bulk");

                    /* Here we initialize the RestClient. Remember to replace your API Key. */
                    RestClient restClient = new RestClient("https://listen.a2g.io/v1/production/microstream");
                    RestRequest restRequest = new RestRequest(Method.POST);
                    restRequest.AddHeader("x-api-key", "");
                    restRequest.AddJsonBody(allData);
                    restRequest.Timeout = 120000;

                    IRestResponse response = restClient.Execute(restRequest);
                    if (response.IsSuccessful == false)
                    {
                        success = false;
                        Console.WriteLine(response.Content);
                        DoSendEmail($"No se envio el control de prevail al inputstream key  <br> restSharp rechazo la conexion. ", "Grafana - Default Trends Home");
                    }
                

            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al enviar el control de prevail al inputstream key  <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "Grafana - Default Trends Home");
            }
            return success;
        }

        private static bool TransferFailureData(List<Dictionary<string, object>> failures, string key)
        {
            int currentIndex = 0;
            bool success = true;
            try
            {
                while (currentIndex < failures.Count)
                {
                    int count = 100;
                    if (currentIndex + count >= failures.Count)
                    {
                        count = failures.Count - currentIndex;
                    }
                    Console.WriteLine($"{currentIndex} to {currentIndex + count} of {failures.Count}");
                    ICollection<Dictionary<string, object>> sub = failures.GetRange(currentIndex, count);
                    currentIndex = currentIndex + count;
                    Dictionary<string, string> allData = new Dictionary<string, string>();
                    allData.Add("IKEY", key);
                    allData.Add("Data", JsonConvert.SerializeObject(sub));
                    allData.Add("Mode", "Bulk");

                    /* Here we initialize the RestClient. Remember to replace your API Key. */
                    RestClient restClient = new RestClient("https://listen.a2g.io/v1/production/microstream");
                    RestRequest restRequest = new RestRequest(Method.POST);
                    restRequest.AddHeader("x-api-key", "");
                    restRequest.AddJsonBody(allData);
                    restRequest.Timeout = 120000;

                    IRestResponse response = restClient.Execute(restRequest);
                    if (response.IsSuccessful == false)
                    {
                        success = false;
                        Console.WriteLine(response.Content);
                    }
                }

            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al enviar la data <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "Grafana - Default Trends Home");
            }
            return success;
        }

        private static bool TransferDrillData(List<Dictionary<string, string>> drills, string key)
        {
            using (StreamWriter file = File.CreateText(@"c:\drills.json"))
            {
                Newtonsoft.Json.JsonSerializer ser = new Newtonsoft.Json.JsonSerializer();
                ser.Serialize(file, drills);
            }

            int currentIndex = 0;
            bool success = true;
            try
            {
                while (currentIndex < drills.Count)
                {
                    int count = 100;
                    if (currentIndex + 100 >= drills.Count)
                    {
                        count = drills.Count - currentIndex;
                    }
                    Console.WriteLine($"{currentIndex} to {currentIndex + count} of {drills.Count}");
                    ICollection< Dictionary<string, string>> sub = drills.GetRange(currentIndex, count);
                    currentIndex = currentIndex + count;
                    Dictionary<string, string> allData = new Dictionary<string, string>();
                    allData.Add("IKEY", key);
                    allData.Add("Data", JsonConvert.SerializeObject(sub));
                    allData.Add("Mode", "Bulk");

                    /* Here we initialize the RestClient. Remember to replace your API Key. */
                    RestClient restClient = new RestClient("https://listen.a2g.io/v1/production/microstream");
                    RestRequest restRequest = new RestRequest(Method.POST);
                    restRequest.AddHeader("x-api-key", "");
                    restRequest.AddJsonBody(allData);
                    restRequest.Timeout = 120000;

                    IRestResponse response = restClient.Execute(restRequest);
                    if (response.IsSuccessful == false)
                    {
                        success = false;
                        Console.WriteLine(response.Content);
                    }
                }

            }
            catch (Exception ex)
            {
                DoSendEmail($"Ocurrió un error al enviar la data <br> Exception: {ex.StackTrace} <br> Exception msg: {ex.Message}", "Grafana - Default Trends Home");
            }
            return success;
        }

        private static ControlPrevail ControlPrevailBuilder(string process, bool wasSuccess, string processCode, int data, float processTime,
                                                            string msg)
        {
            DateTime date = DateTime.UtcNow;
            return new ControlPrevail
            {
                ProcessType = process,
                WasSuccessfull = wasSuccess,
                ProcessCode = processCode,
                DataCount = data,
                ProcessTime = processTime,
                Message = msg,
                Day = new DateTime(date.Year, date.Month, date.Day,0, 0, 0, 0, DateTimeKind.Utc),
                DayHour = new DateTime(date.Year, date.Month, date.Day, date.Hour, 0, 0, 0, DateTimeKind.Utc),
                TimeStamp = date   
            };
        }

        private static bool TransferPrevailData(List<DataPoint> prevailData)
        {
            int currentIndex = 0;
            bool success = true;
            while (currentIndex < prevailData.Count)
            {
                int count = 1000;
                if (currentIndex + 1000 >= prevailData.Count)
                {
                    count = prevailData.Count - currentIndex;
                }
                Console.WriteLine($"{currentIndex} to {currentIndex + count} of {prevailData.Count}");
                ICollection<DataPoint> sub = prevailData.GetRange(currentIndex, count);
                currentIndex = currentIndex + count;
                Dictionary<string, string> allData = new Dictionary<string, string>();
                allData.Add("IKEY", "");
                allData.Add("Data", JsonConvert.SerializeObject(sub));
                allData.Add("Mode", "Bulk");

                /* Here we initialize the RestClient. Remember to replace your API Key. */
                RestClient restClient = new RestClient("https://listen.a2g.io/v1/production/microstream");
                RestRequest restRequest = new RestRequest(Method.POST);
                restRequest.AddHeader("x-api-key", "");
                restRequest.AddJsonBody(allData);
                restRequest.Timeout = 120000;

                IRestResponse response = restClient.Execute(restRequest);
                if (response.IsSuccessful == false)
                {
                    success = false;
                    Console.WriteLine(response.Content);
                }
            }

            return success;
        }

        public static bool DoSendEmail(string body, string window)
        {           
            return true;
        }
    }
}
