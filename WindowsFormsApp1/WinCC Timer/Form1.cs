﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using WinCC_Timer.Properties;
using Interoperability;
using System.Diagnostics;
using System.Threading;
using System.Management;

namespace WinCC_Timer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("msvcrt.dll")]
        private static extern int Memcmp(IntPtr b1, IntPtr b2, long count);

        public DataTable dataTable = new DataTable();
        public string logName = "\\Screen.logger";
        public string pdlrtLogName = "\\pdlrt.logger";
        public string scriptLogName = "\\script.logger";

        public bool endFlag = false;
        public string currentPage = "";

        private void Button1_Click(object sender, EventArgs e)
        {
            Thread.Sleep(5000);
            RunTasks();
        }

        private void Button2_Click(object sender, EventArgs e) => ManipulateWinCCPrograms();

        private void RunTasks()
        {
            #region ParallelTasks
            // Perform tasks in parallel
            Parallel.Invoke(

                () =>
                {
                    Console.WriteLine("Begin first task...");
                    NavigateHMIMenu();
                },  // close first Action

                () =>
                {
                    Console.WriteLine("Begin second task...");
                    GetProcessCPUUsage("PdlRt", pdlrtLogName);
                }, //close second Action

                () =>
                {
                    Console.WriteLine("Begin third task...");
                    GetProcessCPUUsage("script", scriptLogName);
                } //close third Action

            ); //close parallel.invoke

            Console.WriteLine("Returned from Parallel.Invoke");
            #endregion
        }

        private void UpdatePageLabel()
        {
            while (endFlag == false)
            {
                listBox1.Items.Add(currentPage);
                listBox1.Refresh();

                new System.Threading.ManualResetEvent(false).WaitOne(50);
            };
        }

        private void NavigateHMIMenu()
        {
            LogToFile("Starting menu navigation", logName);

            List<MenuRow> lists = GetMenuData(/*(refIdBox.Text*/);

            LogToFile(lists.Count + " menu items", logName);

            var tier1 = lists.Where(c => c.Layer == "1").ToList();

            IntPtr rt = PInvokeLibrary.FindWindow("PDLRTisAliveAndWaitsForYou", "WinCC-Runtime - ");
            if (rt == IntPtr.Zero)
            {
                var msg = "Please start the WinCC Graphics Runtime!";
                listBox1.Items.Add(msg);
                listBox1.Refresh();

                LogToFile(msg, logName);
                return;
            }
            //Bitmap cmp = (Bitmap)Resources.ResourceManager.GetObject("s");
            var singleHeight = 25;

            //"General is 76 * 30
            int x = singleHeight / 2;
            int y = 0;
            foreach (MenuRow m in tier1)
            {
                Size size = GetTextSize(m.Caption, "Arial", 10.0f);

                x += size.Width / 2;
                y = 15;
                y += 28;

                var ChildrenTier1 = lists.Where(c => c.ParentId == m.ID).ToList();
                var tier1Width = GetMenuDropWidth(ChildrenTier1);
                foreach (var tier2 in ChildrenTier1)
                {
                    if (tier2.Pdl != "" && tier2.Pdl != "''")
                    {
                        ClickInWindowAtXY(rt, x, 15, 1); Thread.Sleep(500); //expand tier1 menu
                        LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                        ClickInWindowAtXY(rt, x, y, 1); Thread.Sleep(5000); //open tier2 page
                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                        LogToFile(tier2.Pdl, logName);
                        currentPage = tier2.Pdl;
                        //_ = FindObjectInHMI(cmp);
                    }
                    else
                    {
                        var ChildrenTier2 = lists.Where(c => c.ParentId == tier2.ID);
                        ClickInWindowAtXY(rt, GetMenuDropWidth(ChildrenTier2), y, 1); Thread.Sleep(500); //expand tier2 menu
                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);

                        int yTier3 = y;
                        int maxWidthTier2 = GetMenuDropWidth(ChildrenTier2);
                        LogToFile(maxWidthTier2 + " is maxwidthtier2", logName);
                        foreach (var tier3 in ChildrenTier2)
                        {
                            int xTier3 = x + tier1Width; // + longest element in tier2's width + some 40 pixels
                            if (tier3.Pdl != "" && tier3.Pdl != "''")
                            {
                                ClickInWindowAtXY(rt, x, 15, 1); Thread.Sleep(500); //expand tier1 menu
                                LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                ClickInWindowAtXY(rt, x, y, 1); Thread.Sleep(500); //expand tier2 menu or open page
                                LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                                ClickInWindowAtXY(rt, xTier3, yTier3, 1); Thread.Sleep(5000); //expand tier2 menu or open page
                                LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);

                                LogToFile(tier3.Pdl, logName);

                                currentPage = tier3.Pdl;

                                //_ = FindObjectInHMI(cmp);
                            }
                            else
                            {
                                //there is also a tier4....
                                var ChildrenTier3 = lists.Where(c => c.ParentId == tier3.ID);
                                ClickInWindowAtXY(rt, x, y, 1); Thread.Sleep(500); //expand tier2 menu or open page

                                int yTier4 = yTier3;
                                foreach (var tier4 in ChildrenTier3)
                                {
                                    int xTier4 = xTier3 + GetMenuDropWidth(ChildrenTier3) + 20; // + longest element in tier2's width + some 40 pixels
                                    if (tier4.Pdl != "" && tier4.Pdl != "''")
                                    {
                                        ClickInWindowAtXY(rt, x, 15, 1); Thread.Sleep(500); //expand tier1 menu
                                        LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                        ClickInWindowAtXY(rt, x, y, 1); Thread.Sleep(500); //expand tier2 menu or open page
                                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                                        ClickInWindowAtXY(rt, xTier3, yTier3, 1); Thread.Sleep(500); //expand tier2 menu or open page
                                        LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);
                                        ClickInWindowAtXY(rt, xTier4, yTier4, 1); Thread.Sleep(5000); //expand tier2 menu or open page
                                        LogToFile("For " + tier4.Caption + " expand menu, clicked at " + xTier4 + " x, " + yTier4 + " y", logName);
                                        LogToFile(tier4.Pdl, logName);

                                        currentPage = tier4.Pdl;
                                        //_ = FindObjectInHMI(cmp);
                                    }
                                    else
                                    {
                                        //no tier 5 thank god
                                    }
                                    yTier4 += singleHeight; //if 26 here, then it becomes too much
                                }
                            }
                            yTier3 += singleHeight;
                        }
                    }
                    y += singleHeight; //increment tier2 menu or open page
                }

                x += size.Width / 2;
                x += singleHeight;

                //bool v = FindObjectInHMI(cmp);
                //still only almost, gets too offset in the end
            }
            endFlag = true;
        }

        private int GetMenuDropWidth(IEnumerable<MenuRow> ChildrenTier2)
        {
            int maxWidthTier2 = 0;
            foreach (var z in ChildrenTier2)
            {
                var w = GetTextSize(z.Caption, "Arial", 10.0f).Width;
                if (w > maxWidthTier2)
                {
                    LogToFile(z.Caption + " is the caption for which we check width", logName);
                    maxWidthTier2 = w;
                }
            }

            LogToFile(maxWidthTier2 + " is x width for childrentier2", logName);
            return maxWidthTier2 + 40;
        }

        private Size GetTextSize(string text, string fontName, Single fontSize)
        {
            Font font = new Font(fontName, fontSize, FontStyle.Regular);
            Size size = TextRenderer.MeasureText(text, font);
            size.Width -= 8; //stupid calculation
            LogToFile(size.Width.ToString() + "width, " + size.Height.ToString() + "height, " + text, logName);
            //26 height for submenus, 30 height for the tier1 menu
            return size;
        }

        private void ClickInWindowAtXY(IntPtr handle, int? x, int? y, int repeat)
        {
            for (int i = 0; i < repeat; i++)
            {
                PInvokeLibrary.SetForegroundWindow(handle);

                MouseOperations.SetCursorPosition(x.Value, y.Value); //have to use the found minus/plus coordinates here
                MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
            }
        }

        private List<MenuRow> GetMenuData(/*string id*/)
        {
            #region sql
            ////sql connection to fill data 
            //string connectionString = "Data Source=" + sqlPathBox.Text + "\\WINCC;Initial Catalog=SMS_RTDesign;Integrated Security=SSPI";
            //string query = "SELECT * FROM RT_Menu where RefId = " + id;
            //SqlConnection cnn = new SqlConnection(connectionString);
            //try
            //{
            //    cnn.Open();
            //    //MessageBox.Show("Connection Open ! ");
            //    SqlCommand cmd = new SqlCommand(query, cnn);
            //    SqlDataAdapter da = new SqlDataAdapter(cmd);
            //    da.Fill(dataTable);
            //    cnn.Close();
            //}
            //catch (Exception ex)
            //{
            //    LogToFile("Can not open connection ! " + ex.Message, logName);
            //    return null;
            //}
            #endregion

            string sqlFile = Application.StartupPath + @"\" + textBox1.Text;

            var fileInfo = new FileInfo(sqlFile);

            if (!fileInfo.Exists)
            {

                return new List<MenuRow>();
            }

            //var g = new grafexe.Application().ApplicationDataPath;

            //INSERT INTO[SMS_RTDesign].[dbo].[RT_Menu](ID, RefId, Layer, Pos, Parentid, LCID, Caption, Flags, Pdl) VALUES (1000000, 1, 1, 1, 0, 1033, N'General', 3, '')
            //INSERT INTO[SMS_RTDesign].[dbo].[RT_Menu](ID, RefId, Layer, Pos, Parentid, LCID, Caption, Flags, Pdl) VALUES (1010000, 1, 2, 1, 1000000, 1033, N'Alarms', 1, '@sms_w_Alarmlist')

            var textData = File.ReadAllLines(sqlFile);
            var manipData = new List<string>();
            for (int i = 0; i < textData.Length; i++)
            {
                string line = textData[i];
                if (line.IndexOf("VALUES (") != -1)
                    manipData.Add(line.Substring(
                        line.IndexOf("VALUES (") + "VALUES (".Length,
                        line.Length - 1 - line.IndexOf("VALUES (") - "VALUES (".Length
                        ));
            }

            List<MenuRow> myData = new List<MenuRow>();

            for (int i = 0; i < manipData.Count; i++)
            {
                var row = manipData[i].Split(Convert.ToChar(","));
                myData.Add(new MenuRow()
                {
                    ID = row?.ElementAt(0).ToString(),
                    RefId = row?.ElementAt(1).ToString(),
                    Layer = row?.ElementAt(2).ToString(),
                    Pos = row?.ElementAt(3).ToString(),
                    ParentId = row?.ElementAt(4).ToString(),
                    LCID = row?.ElementAt(5).ToString(),
                    Caption = row?.ElementAt(6).ToString().Replace("N'", "").Replace("'", ""),
                    Flags = row?.ElementAt(7).ToString(),
                    Pdl = row?.ElementAt(8).ToString(),
                    //Parameter = row?.ElementAt(9)?.ToString(),
                });
                Console.WriteLine("added row");
            }

            return myData;
        }

        private bool FindObjectInHMI(Bitmap cmp)
        {
            DateTime start = DateTime.UtcNow;
            Bitmap bmp = TakeScreenShot(1681, 1004);
            bool t = CompareMemCmp(bmp, cmp);
            LogToFile((DateTime.UtcNow - start).TotalMilliseconds.ToString() + "ms", logName);
            return t;
        }

        public static bool CompareMemCmp(Bitmap b1, Bitmap b2)
        {
            if ((b1 == null) != (b2 == null)) return false;
            if (b1.Size != b2.Size) return false;

            var bd1 = b1.LockBits(new Rectangle(new Point(0, 0), b1.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var bd2 = b2.LockBits(new Rectangle(new Point(0, 0), b2.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                IntPtr bd1scan0 = bd1.Scan0;
                IntPtr bd2scan0 = bd2.Scan0;

                int stride = bd1.Stride;
                int len = stride * b1.Height;

                return Memcmp(bd1scan0, bd2scan0, len) == 0;
            }
            finally
            {
                b1.UnlockBits(bd1);
                b2.UnlockBits(bd2);
            }
        }

        private void LogToFile(string content, string fname)
        {
            using (var fileWriter = new StreamWriter(Application.StartupPath + fname, true))
            {
                DateTime date = DateTime.UtcNow;
                fileWriter.WriteLine(date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content);
                fileWriter.Close();
            }
        }

        private Bitmap TakeScreenShot(int left, int top)
        {
            int width = 5;
            int height = 5;
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                Size s = new Size()
                {
                    Width = width, //Screen.PrimaryScreen.Bounds.Size.Width - width,
                    Height = height //Screen.PrimaryScreen.Bounds.Size.Height - height
                };
                g.CopyFromScreen(left, top, 0, 0, s);
                //bmp.Save("s.png");  // saves the image
            }
            return bmp;
        }

        private static void ManipulateWinCCPrograms()
        {
            #region grafexe
            grafexe.Application g = new grafexe.Application();

            var file = @"C:\Project\sdib_tcm_clt\GraCS\TCM#01-01-01_n_#TCM-OverviewTCM.pdl";
            var grf = g.Documents.Open(file, grafexe.HMIOpenDocumentType.hmiOpenDocumentTypeVisible);

            g.DisableVBAEvents = false;

            var myobjects = grf.HMIObjects.Find("HMIRectangle");
            var count = myobjects.Count;

            grafexe.HMIObjects objs = grf.HMIObjects;
            foreach (grafexe.HMIObject obj in objs)
            {
                var t = obj.ObjectName;
            }
            #endregion

            #region runtime
            CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();

            var hmiRtProjectLib = new CCHMIRTPROJECTLib.HMIProject();
            var hmiRtGraphics = new CCHMIRTGRAPHICS.HMIRTGraphics();
            //var agentComp = new CCAGENTLib.CCAgentComp().;
            var configToolImporter = new CCCONFIGTOOLIMPORTERLib.CCConfigToolImporter();
            var downloadLib = new CCDOWNLOADLib.InitDownload();
            downloadLib.Initialize(CCDOWNLOADLib.enumWinCCMode.WCM_RT, CCDOWNLOADLib.enumClientType.CLT_WINCC);

            var guiTools = new CCGUITOOLSLib.CCBalloon();
            guiTools.HideBalloon();

            #endregion

            #region configStudio
            var tag = new CCConfigStudio.Application();
            tag.Editors[1].FileOpen("");
            #endregion
        }

        private bool GetProcessCPUUsage(string process, string log)
        {
            LogToFile("Starting cpu usage gathering", logName);

            var name = string.Empty;
            var perc = new List<float>();
            var measTime = new List<string>();
            var atPagesList = new List<string>();
            var datetimes = new List<DateTime>();
            var procs = Environment.ProcessorCount;

            foreach (var proc in Process.GetProcesses())
            {
                if (proc.ProcessName.StartsWith(process))
                {
                    name = proc.ProcessName;
                    proc.StartInfo.RedirectStandardOutput = true;
                    LogToFile("Found process with name " + name, logName);
                }
            }

            PerformanceCounter cpu = new PerformanceCounter("Process", "% Processor Time", name, true);

            try
            {
                while (endFlag == false)
                {
                    DateTime date = DateTime.UtcNow;
                    datetimes.Add(date);
                    perc.Add(cpu.NextValue() / procs);
                    measTime.Add(date.Hour + ":" + date.Minute + ":" + date.Second + "." + date.Millisecond);
                    atPagesList.Add(currentPage);
                    Thread.Sleep(50);
                }
            }
            catch (Exception exc)
            {
                LogToFile(exc.Message, logName);
            }

            for (int i = 0; i < perc.Count; i++)
            {
                LogToFile(measTime[i] + "," + perc[i].ToString() + "," + atPagesList[i], log);
            }

            try
            {
                ProcessGatheredCpuUsageData(perc, atPagesList, datetimes, "current");
            }
            catch (Exception exc)
            {
                LogToFile(exc.Message, logName);
            }

            return true;
        }

        private void ProcessGatheredCpuUsageData(List<float> perc, List<string> atPagesList, List<DateTime> datetimes, string currentCalc)
        {
            var PageCpuUsageList = new List<PageCpuTime>();
            var PageLoadTimes = new List<PageTime>();
            for (int i = 0; i < perc.Count; i++)
            {
                PageCpuUsageList.Add(new PageCpuTime()
                {
                    cpu = perc[i],
                    page = atPagesList[i],
                    timestamp = datetimes[i]
                });
            }

            var pagesNavigated = atPagesList.Distinct().Where(c => c != "").ToList();
            foreach (string p in pagesNavigated)
            {
                List<PageCpuTime> currentPageData = new List<PageCpuTime>();
                List<PageCpuTime> tempList = PageCpuUsageList.Where(c => c.page == p).OrderBy(c => c.timestamp).ToList();
                var startTime = tempList.Select(c => c.timestamp).Min();

                List<List<PageCpuTime>> nonZeroGroups = new List<List<PageCpuTime>>
                {
                    new List<PageCpuTime>()
                };
                foreach (var c in tempList)
                {
                    if (c.cpu > 0)
                    {
                        nonZeroGroups.Last().Add(c);
                    }
                    else
                    {
                        nonZeroGroups.Add(new List<PageCpuTime>());
                    }
                }

                nonZeroGroups = nonZeroGroups.Where(c => c.Count > 0).ToList();

                foreach (var c in nonZeroGroups.ToList())
                {
                    PageCpuTime prim = c.OrderBy(x => x.timestamp).FirstOrDefault();
                    if ((prim.timestamp - startTime).TotalSeconds > 2)
                        nonZeroGroups.Remove(c);
                } //remove groups that start later than 2 seconds

                var largestNonZero = nonZeroGroups.OrderByDescending(c => c.Count()).ElementAt(0).OrderBy(c => c.timestamp).ToList();
                double loadingTime = (largestNonZero.LastOrDefault().timestamp - startTime).TotalMilliseconds;

                PageLoadTimes.Add(new PageTime()
                {
                    load = loadingTime,
                    page = p
                });

                LogToFile(PageLoadTimes.LastOrDefault().page + "," + PageLoadTimes.LastOrDefault().load + " ms", "\\timerData_" + currentCalc + ".logger");
            }
        }

        public class PageTime
        {
            public string page;
            public double load;
        }

        public class PageCpuTime
        {
            public string page;
            public float cpu;
            public DateTime timestamp;
        }

        public class MenuRow
        {
            public string ID;
            public string RefId;
            public string Layer;
            public string Pos;
            public string LCID;
            public string ParentId;
            public string Caption;
            public string Flags;
            public string Pdl;
            public string Parameter;
        }

        #region TestingWMI
        private static void WMIQueryForCPUUsage()
        {
            //Get CPU usage values using a WMI query
            for (int i = 0; i < 3; i++)
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("select * from Win32_PerfFormattedData_PerfProc_Process WHERE Name = 'Taskmgr'");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    Console.WriteLine("Name: {0}", queryObj["Name"]);
                    Console.WriteLine("PercentProcessorTime: {0}", queryObj["PercentProcessorTime"]);

                }
            }
        }

        private static int lineCount = 0;
        private static StringBuilder output = new StringBuilder();
        public static void ReadProcessOutputTest()
        {
            Process process = Process.GetProcessesByName("Taskmgr")[0];
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.OutputDataReceived += new DataReceivedEventHandler((sender, e) =>
            {
                // Prepend line numbers to each line of the output.
                if (!String.IsNullOrEmpty(e.Data))
                {
                    lineCount++;
                    output.Append("\n[" + lineCount + "]: " + e.Data);
                }
            });

            // Asynchronously read the standard output of the spawned process.
            // This raises OutputDataReceived events for each line of output.
            process.BeginOutputReadLine();
            process.WaitForExit();

            // Write the redirected output to this application's window.
            Console.WriteLine(output);

            process.WaitForExit();
            process.Close();

            Console.WriteLine("\n\nPress any key to exit.");
            Console.ReadLine();
        }
        #endregion

        private void Button3_Click(object sender, EventArgs e)
        {
            var file = textBox2.Text;

            var dataFile = new FileInfo(Application.StartupPath + @"\" + file + ".logger");
            if (!dataFile.Exists) return;

            var textData = File.ReadAllLines(dataFile.FullName);
            var perc = new List<float>();
            var atPagesList = new List<string>();
            var datetimes = new List<DateTime>();

            for (int i = 0; i < textData.Length; i++)
            {
                var data = textData[i].Split(Convert.ToChar(" "))[3];
                var line = data.Split(Convert.ToChar(","));
                datetimes.Add(DateTime.Parse(line[0]));
                perc.Add(float.Parse(line[1]));
                atPagesList.Add(line[2]);
            }

            ProcessGatheredCpuUsageData(perc, atPagesList, datetimes, file);
        }

        private void Button4_Click(object sender, EventArgs e)
        {

        }
    }
}