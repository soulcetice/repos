﻿using CCHMIRUNTIME;
using CommonInterops;
using grafexe;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinCC_Timer.Properties;
using WindowsUtilities;
using RuntimeUtilities;

namespace WinCC_Timer
{
    public partial class Form1 : Form
    {
        public bool collapsed = true;

        public Form1()
        {
            InitializeComponent();

            InitTreeView();

            SetTooltips();

            FileInfo sqlFiles = FindSQLFile();
            if (sqlFiles != null)
                textBox1.Text = sqlFiles.Name;
        }

        private FileInfo FindSQLFile()
        {
            var d = new DirectoryInfo(System.Windows.Forms.Application.StartupPath);
            var sqlFiles = Directory.GetFiles(d.FullName, "*.sql", SearchOption.TopDirectoryOnly).ToList();
            if (sqlFiles.Count > 0)
                return new FileInfo(sqlFiles.FirstOrDefault());
            else
            {
                listBox1.Items.Add("No .sql file was found in the folder!");
                return null;
            }
        }

        private void UpdateFileDate()
        {
            SetDateString();
            Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\" + formattedDate);
        }

        private void SetDateString()
        {
            DateTime date = DateTime.Now;
            var month = date.Month < 10 ? "0" + date.Month.ToString() : date.Month.ToString();
            var day = date.Day < 10 ? "0" + date.Day.ToString() : date.Day.ToString();
            var hour = date.Hour < 10 ? "0" + date.Hour.ToString() : date.Hour.ToString();
            var minute = date.Minute < 10 ? "0" + date.Minute.ToString() : date.Minute.ToString();
            var second = date.Second < 10 ? "0" + date.Second.ToString() : date.Second.ToString();
            formattedDate = date.Year.ToString().Substring(2) + month + day + hour + minute + second;
        }

        private void SetTooltips()
        {
            SetTooltip("Set the number of measurements to take", numberOfMeasurements);
            SetTooltip("Time interval between measurements", timeInterval); // (around 2 * pagesNo [minutes])
            SetTooltip("Number of hours the runs will take", hoursToRun);
            SetTooltip("Run the set number of measurements at the set minutes interval!", button1);
            SetTooltip("Include screenshots in the calculations' folders", checkBox4);
            SetTooltip(@"Take screenshot after 1.5 seconds (found in the ""Screenshots"" Folder)", snapBox);
            SetTooltip("Expand/Collapse all tree elements", checkBox1);
            SetTooltip("Select/Deselect all tree elements", checkBox2);
            SetTooltip("Calculate averages on the gathered datasets using SD calculation", button5);
            SetTooltip("Menu SQL file reflecting current menu state on the machine", textBox1);
            SetTooltip("Folder containing calculations to use for generating averaged measurements", textBox3);
            SetTooltip("Show/Hide calculations panel", checkBox3);
            SetTooltip(@"Export viewable data as a "".csv."" file", export);
            SetTooltip("Regenerate timer data files in the calculations folder (from the pdlrt files)", button3);
            SetTooltip("Show data for the selected pages in an excel workbook (in progress)", button6);
            SetTooltip("Feedback messages", listBox1);
        }

        private static void SetTooltip(string msg, dynamic obj)
        {
            var tooltip = new System.Windows.Forms.ToolTip();
            tooltip.SetToolTip(obj, msg);
            tooltip.InitialDelay = 50;
            tooltip.ReshowDelay = 100;
            tooltip.UseFading = true;
        }

        private void InitTreeView()
        {

            ////
            //// This is the first node in the view.
            ////
            TreeNode treeNode = new TreeNode("Windows");
            //treeView1.Nodes.Add(treeNode);
            ////
            //// Another node following the first node.
            ////
            //treeNode = new TreeNode("Linux");
            //treeView1.Nodes.Add(treeNode);
            ////
            //// Create two child nodes and put them in an array.
            //// ... Add the third node, and specify these as its children.
            ////
            TreeNode node2 = new TreeNode("C#");
            TreeNode node3 = new TreeNode("VB.NET");
            TreeNode[] array = new TreeNode[] { node2, node3 };
            ////
            //// Final node.
            ////
            //treeNode = new TreeNode("Dot Net Perls", array);
            //treeView1.Nodes.Add(treeNode);
            List<MenuRow> lists = GetMenuData(/*(refIdBox.Text*/);

            var tier1 = lists.Where(c => c.Layer == "1").ToList();

            foreach (MenuRow m in tier1)
            {
                treeNode = new TreeNode(m.Caption);
                var ChildrenTier1 = lists.Where(c => c.ParentId == m.ID).ToList();
                array = new TreeNode[ChildrenTier1.Count];

                for (int i = 0; i < ChildrenTier1.Count; i++)
                {
                    MenuRow child1 = ChildrenTier1[i];
                    var ChildrenTier2 = lists.Where(c => c.ParentId == child1.ID).ToList();
                    TreeNode[] array1 = new TreeNode[ChildrenTier2.Count];

                    for (int i1 = 0; i1 < ChildrenTier2.Count; i1++)
                    {
                        MenuRow child2 = ChildrenTier2[i1];
                        var ChildrenTier3 = lists.Where(c => c.ParentId == child2.ID).ToList();
                        TreeNode[] array2 = new TreeNode[ChildrenTier3.Count];

                        for (int i2 = 0; i2 < ChildrenTier3.Count; i2++)
                        {
                            MenuRow child3 = ChildrenTier3[i2];
                            var ChildrenTier4 = lists.Where(c => c.ParentId == child3.ID).ToList();
                            TreeNode[] array3 = new TreeNode[ChildrenTier4.Count];

                            for (int i3 = 0; i3 < ChildrenTier4.Count; i3++)
                            {
                                MenuRow child4 = ChildrenTier4[i3];
                                array3[i3] = new TreeNode(child4.Caption);
                            }

                            array2[i2] = new TreeNode(child3.Caption, array3);
                        }

                        array1[i1] = new TreeNode(child2.Caption, array2);
                    }

                    array[i] = new TreeNode(child1.Caption, array1);
                }
                treeNode = new TreeNode(m.Caption, array);
                treeView1.Nodes.Add(treeNode);
            }

            treeView1.CheckBoxes = true;

            SelectOrDeselectAllTree(false);
        }

        [DllImport("msvcrt.dll")]
        private static extern int Memcmp(IntPtr b1, IntPtr b2, long count);

        public DataTable dataTable = new DataTable();
        public string logName = "\\Screen.logger";

        public bool endFlag = false;
        public string currentPage = "";
        public string currentActiveScreen = "";
        public string formattedDate = "";
        public bool firstRunHasEnded = false;

        private void Button1_Click(object sender, EventArgs e)
        {
            FileInfo sqlFiles = FindSQLFile();
            if (sqlFiles == null)
            {
                return;
            }
            for (var i = 0; i < Int32.Parse(numberOfMeasurements.Text); i++)
            {
                endFlag = false;
                Thread.Sleep(4000);
                RunTasksForTiming();
            }
        }

        private void Button2_Click(object sender, EventArgs e) => ManipulateWinCCPrograms();

        private void RunTasksForTiming()
        {
            UpdateFileDate();
            #region ParallelTasks
            // Perform tasks in parallel
            Parallel.Invoke(

                () =>
                {
                    Console.WriteLine("Begin first task...");
                    NavigateHMIMenu(scour: false);
                },  // close first Action

                () =>
                {
                    Console.WriteLine("Begin second task...");
                    GatherProcessCPUUsage("PdlRt", "\\pdlrt_" + formattedDate + ".logger");
                }, //close second Action

                () =>
                {
                    Console.WriteLine("Begin third task...");
                    CloseSoftwareWarnings();
                } //close third Action

            ); //close parallel.invoke

            Console.WriteLine("Returned from Parallel.Invoke");
            #endregion
        }

        private void RunTasksForScreenshots()
        {
            UpdateFileDate();
            #region ParallelTasks
            // Perform tasks in parallel
            Parallel.Invoke(

                () =>
                {
                    Console.WriteLine("Begin first task...");
                    NavigateHMIMenu(scour: true);
                },  // close first Action

                () =>
                {
                    Console.WriteLine("Begin third task...");
                    CloseSoftwareWarnings();
                } //close third Action

            ); //close parallel.invoke

            Console.WriteLine("Returned from Parallel.Invoke");
            #endregion
        }

        private void CloseSoftwareWarnings()
        {
            var anyPopupClass = "#32770"; //usually any popup
            while (endFlag == false)
            {
                IntPtr wrg = PInvokeLibrary.FindWindow(anyPopupClass, "WinCC Information");
                if (wrg != IntPtr.Zero)
                {
                    PInvokeLibrary.SetForegroundWindow(wrg);

                    IntPtr DlButtonHandle = PInvokeLibrary.FindWindowEx(wrg, IntPtr.Zero, "Button", "OK");
                    if (DlButtonHandle != IntPtr.Zero)
                    {
                        PInvokeLibrary.SendMessage(DlButtonHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                    }
                }
                Thread.Sleep(100);
            }
        }

        private List<string> selectedNodes = new List<string>();

        public void GetCheckedNodes(TreeNodeCollection nodes)
        {
            foreach (System.Windows.Forms.TreeNode aNode in nodes)
            {
                if (!aNode.Checked)
                    continue;

                selectedNodes.Add(aNode.Text);

                if (aNode.Nodes.Count != 0)
                    GetCheckedNodes(aNode.Nodes);
            }
        }

        private void NavigateHMIMenu(bool scour = false)
        {
            LogToFile("Starting menu navigation", logName);

            List<MenuRow> lists = GetMenuData(/*(refIdBox.Text*/);

            LogToFile(lists.Count + " menu items", logName);

            GetCheckedNodes(treeView1.Nodes);

            var tier1 = lists.Where(c => c.Layer == "1").ToList();
            IntPtr rt = GetHandle("PDLRTisAliveAndWaitsForYou", "WinCC-Runtime - ");
            if (rt == IntPtr.Zero)
            {
                return;
            }

            if (scour)
            {
                GetRuntimeInfo("HMIButton", out List<ObjectData> extractedData, "@");
                FindOpenCloseFlyouts(rt, extractedData); //first run, run only once this way
            }

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
                        if (!selectedNodes.Contains(tier2.Pdl))
                        {
                            RobustClick(rt, x, 15, 1); Thread.Sleep(1000); //expand tier1 menu
                            LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);

                            RobustClick(rt, x, y, 1); currentPage = tier2.Pdl;
                            Thread.Sleep(1500);
                            if (snapBox.Checked && !firstRunHasEnded) //only fpr snapshot to ensure page has loaded after tolerance time
                                ScreenshotAndSave(true);
                            Thread.Sleep(5000); //open tier2 page

                            LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                            LogToFile(tier2.Pdl, logName);

                            if (checkBox4.Checked)
                                ScreenshotAndSave();

                            if (scour)
                                ScourPage();
                        }
                    }
                    else
                    {
                        var ChildrenTier2 = lists.Where(c => c.ParentId == tier2.ID);
                        RobustClick(rt, GetMenuDropWidth(ChildrenTier2), y, 1); Thread.Sleep(1000); //expand tier2 menu
                        LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);

                        int yTier3 = y;
                        int maxWidthTier2 = GetMenuDropWidth(ChildrenTier2);
                        LogToFile(maxWidthTier2 + " is maxwidthtier2", logName);
                        foreach (var tier3 in ChildrenTier2)
                        {
                            int xTier3 = x + tier1Width; // + longest element in tier2's width + some 40 pixels
                            if (tier3.Pdl != "" && tier3.Pdl != "''")
                            {
                                if (!selectedNodes.Contains(tier3.Pdl))
                                {
                                    RobustClick(rt, x, 15, 1); Thread.Sleep(1000); //expand tier1 menu
                                    LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                    RobustClick(rt, x, y, 1); Thread.Sleep(1000); //expand tier2 menu or open page
                                    LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);

                                    RobustClick(rt, xTier3, yTier3, 1); currentPage = tier3.Pdl;
                                    Thread.Sleep(1500);
                                    if (snapBox.Checked && !firstRunHasEnded) //only fpr snapshot to ensure page has loaded after tolerance time
                                        ScreenshotAndSave(true);
                                    Thread.Sleep(5000); //expand tier2 menu or open page

                                    LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);
                                    LogToFile(tier3.Pdl, logName);

                                    if (checkBox4.Checked)
                                        ScreenshotAndSave();

                                    if (scour)
                                        ScourPage();
                                }
                            }
                            else
                            {
                                //there is also a tier4....
                                var ChildrenTier3 = lists.Where(c => c.ParentId == tier3.ID);
                                RobustClick(rt, x, y, 1); Thread.Sleep(1000); //expand tier2 menu or open page

                                int yTier4 = yTier3;
                                foreach (var tier4 in ChildrenTier3)
                                {
                                    var tier4Width = GetTextSize(tier4.Caption, "Arial", 10.0f).Width / 2 + 20;
                                    LogToFile(tier4.Caption + "'s click position in pixels would be " + tier4Width.ToString(), logName);
                                    int xTier4 = xTier3 + GetMenuDropWidth(ChildrenTier3) + 40; // + longest element in tier2's width + some 40 pixels
                                    if (tier4.Pdl != "" && tier4.Pdl != "''")
                                    {
                                        if (!selectedNodes.Contains(tier4.Pdl))
                                        {
                                            RobustClick(rt, x, 15, 1); Thread.Sleep(1000); //expand tier1 menu
                                            LogToFile("For " + m.Caption + " expand menu, clicked at " + x + " x, " + 15 + " y", logName);
                                            RobustClick(rt, x, y, 1); Thread.Sleep(1000); //expand tier2 menu or open page
                                            LogToFile("For " + tier2.Caption + " expand menu, clicked at " + x + " x, " + y + " y", logName);
                                            RobustClick(rt, xTier3, yTier3, 1); Thread.Sleep(1000); //expand tier2 menu or open page
                                            LogToFile("For " + tier3.Caption + " expand menu, clicked at " + xTier3 + " x, " + yTier3 + " y", logName);

                                            RobustClick(rt, xTier4, yTier4, 1); currentPage = tier4.Pdl;
                                            Thread.Sleep(1500);
                                            if (snapBox.Checked && !firstRunHasEnded) //only fpr snapshot to ensure page has loaded after tolerance time
                                                ScreenshotAndSave(true);
                                            Thread.Sleep(5000); //expand tier2 menu or open page

                                            LogToFile("For " + tier4.Caption + " expand menu, clicked at " + xTier4 + " x, " + yTier4 + " y", logName);
                                            LogToFile(tier4.Pdl, logName);

                                            if (checkBox4.Checked)
                                                ScreenshotAndSave();

                                            if (scour)
                                                ScourPage();
                                        }
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

                //still only almost, gets too offset in the end
            }
            endFlag = true;
            firstRunHasEnded = true;
        }

        private IntPtr GetHandle(string className, string windowName)
        {
            IntPtr rt = PInvokeLibrary.FindWindow(className, windowName); ;
            if (rt == IntPtr.Zero)
            {
                var msg = "Please start the WinCC Graphics Runtime!";
                listBox1.Items.Add(msg);
                listBox1.Refresh();

                LogToFile(msg, logName);
            }

            return rt;
        }

        private void ScreenshotAndSave(bool flatFolder = false, Rectangle rect = new Rectangle())
        {
            if (rect.Width == 0)
            {
                rect.X = 0;
                rect.Y = 0;
                rect.Height = Screen.PrimaryScreen.Bounds.Height;
                rect.Width = Screen.PrimaryScreen.Bounds.Width;
            }

            Bitmap bmp = TakeScreenShot(rect); //0, 0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            string path = "";
            if (!flatFolder)
            {
                path = System.Windows.Forms.Application.StartupPath + "\\" + formattedDate + "\\" + currentPage + ".png";
            }
            else
            {
                if (new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\" + "Screenshots").Exists == false)
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\" + "Screenshots");
                path = System.Windows.Forms.Application.StartupPath + "\\" + "Screenshots" + "\\" + currentPage + "_" + inc + ".png";
            }

            if (new FileInfo(path).Exists)
            {
                inc++;
                path = path.Replace(currentPage + "_" + (inc - 1), currentPage + "_" + inc);
            }

            bmp.Save(path);
        }
        int inc = 0;

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

        private void RobustClick(IntPtr handle, int? x, int? y, int repeat)
        {
            for (int i = 0; i < repeat; i++)
            {
                PInvokeLibrary.SetForegroundWindow(handle);
                ClickAtXY(x, y);
            }
        }

        private static void ClickAtXY(int? x, int? y)
        {
            MouseOperations.SetCursorPosition(x.Value, y.Value); //have to use the found minus/plus coordinates here
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            //Thread.Sleep(500);
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
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

            string sqlFile = System.Windows.Forms.Application.StartupPath + @"\" + textBox1.Text;

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

            var recommendedIntervalMins = myData.Where(c => c.Pdl != "''").Count() / 10 * 1.5;
            timeInterval.Text = recommendedIntervalMins.ToString();
            //numberOfMeasurements.Text = (Math.Floor(Double.Parse(hoursToRun.Text) * 60 / recommendedIntervalMins)).ToString();
            hoursToRun.Text = Math.Round((recommendedIntervalMins * Double.Parse(numberOfMeasurements.Text)) / 60, 1).ToString();

            return myData;

        }

        private void LogToFile(string content, string fname, bool useDate = true)
        {
            using (var fileWriter = new StreamWriter(System.Windows.Forms.Application.StartupPath + fname, true))
            {
                DateTime date = DateTime.UtcNow;
                if (useDate)
                {
                    content = date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content;
                }

                fileWriter.WriteLine(content);
                fileWriter.Close();
            }
        }

        private Bitmap TakeScreenShot(Rectangle rect /*int left, int top, int width, int height*/)
        {
            Bitmap bmp = new Bitmap(rect.Width, rect.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                Size s = new Size()
                {
                    Width = rect.Width, //Screen.PrimaryScreen.Bounds.Size.Width - width,
                    Height = rect.Height //Screen.PrimaryScreen.Bounds.Size.Height - height
                };
                g.CopyFromScreen(rect.X, rect.Y, 0, 0, s);
            }
            return bmp;
        }

        private static void ManipulateWinCCPrograms()
        {
            //#region grafexe
            //grafexe.Application g = new grafexe.Application();

            //var file = @"C:\Project\sdib_tcm_clt\GraCS\TCM#01-01-01_n_#TCM-OverviewTCM.pdl";
            //var grf = g.Documents.Open(file, grafexe.HMIOpenDocumentType.hmiOpenDocumentTypeVisible);

            //g.DisableVBAEvents = false;

            //var myobjects = grf.HMIObjects.Find("HMIRectangle");
            //var count = myobjects.Count;

            //grafexe.HMIObjects objs = grf.HMIObjects;
            //foreach (grafexe.HMIObject obj in objs)
            //{
            //    var t = obj.ObjectName;
            //}
            //#endregion

            //#region runtime
            //CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();

            //var hmiRtProjectLib = new CCHMIRTPROJECTLib.HMIProject();
            //var hmiRtGraphics = new CCHMIRTGRAPHICS.HMIRTGraphics();
            ////var agentComp = new CCAGENTLib.CCAgentComp().;
            //var configToolImporter = new CCCONFIGTOOLIMPORTERLib.CCConfigToolImporter();
            //var downloadLib = new CCDOWNLOADLib.InitDownload();
            //downloadLib.Initialize(CCDOWNLOADLib.enumWinCCMode.WCM_RT, CCDOWNLOADLib.enumClientType.CLT_WINCC);

            //var guiTools = new CCGUITOOLSLib.CCBalloon();
            //guiTools.HideBalloon();

            //#endregion

            //#region configStudio
            var ed = new CCConfigStudio.Application();
            var editors = ed.Editors.Cast<CCConfigStudio.Editor>().ToList();
            //tm[1].FileOpen("");

            CCConfigStudio.Editor tm = editors.FirstOrDefault(c => c.Name == "Tag Management");

            //tm.Select().

            //tm.Select(tree)

            Console.WriteLine("");
            //#endregion
        }

        private bool GatherProcessCPUUsage(string process, string log)
        {
            //LogToFile("Starting cpu usage gathering", logName);

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
                    //LogToFile("Found process with name " + name, logName);
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
                LogToFile(exc.Message, "\\errors.logger");
            }

            for (int i = 0; i < perc.Count; i++)
            {
                LogToFile(measTime[i] + "," + perc[i].ToString() + "," + atPagesList[i], "\\" + formattedDate + "\\" + log);
            }

            try
            {
                ProcessGatheredCpuUsageData(perc, atPagesList, datetimes, formattedDate);
            }
            catch (Exception exc)
            {
                LogToFile(exc.Message, "\\errors.logger");
            }

            return true;
        }

        private void ProcessGatheredCpuUsageData(List<float> perc, List<string> atPagesList, List<DateTime> datetimes, string currentCalc, string directory = "")
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
                List<PageCpuTime> tempList = PageCpuUsageList.Where(c => c.page == p).ToList();
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

                nonZeroGroups = removeIrrelevantCpuGroups(startTime, nonZeroGroups);

                if (nonZeroGroups.Count > 0)
                {
                    var largestNonZero = nonZeroGroups.OrderByDescending(c => c.Count()).ElementAt(0).ToList();

                    var difs = new List<double>();

                    if (largestNonZero.Count > 1)
                    {
                        for (int i = 0; i < largestNonZero.Count - 1; i++)
                        {
                            difs.Add(largestNonZero[i + 1].cpu - largestNonZero[i].cpu);
                        }
                    }

                    //largest three compounded steps is the big spike
                    //which is really the point where the datamanager requests and wincc shows the data
                    //aka the loading time as the eye sees it

                    difs = difs.Select(x => (x < 0 ? 0 : x)).ToList();
                    GetConsecutiveSum(largestNonZero.Select(c => c.cpu).ToArray(), out float maxCompoundSlope, out int index);

                    double loadingTime = (largestNonZero[index].timestamp - startTime).TotalMilliseconds;
                    //not the max but actually the point after the steepest slope we got !

                    var pageTime = new PageTime()
                    {
                        load = loadingTime,
                        page = p
                    };
                    PageLoadTimes.Add(pageTime);
                    LogToFile(pageTime.page + "," + pageTime.load + " ms", directory.Replace(System.Windows.Forms.Application.StartupPath, "") + "\\timerData_" + currentCalc + ".logger");
                }
            }
        }

        private static void GetConsecutiveSum(float[] values, out float maximumSum, out int startIndex)
        {
            maximumSum = 0;
            startIndex = -1;
            for (int i = 0; i < values.Length - 2; i++)
            {
                float prev = 0;
                if (i == 0)
                    prev = 0;
                else
                    prev = values[i - 1];

                float sequenceSum = values[i] + values[i + 1] + values[i + 2] - prev * 3; //- i - 1 to remove start point
                if (sequenceSum > maximumSum)
                {
                    maximumSum = sequenceSum;
                    startIndex = i + 2;
                }
            }

            if (startIndex == -1)
                startIndex = values.Count() - 1;
        }

        private static List<List<PageCpuTime>> removeIrrelevantCpuGroups(DateTime startTime, List<List<PageCpuTime>> nonZeroGroups)
        {
            nonZeroGroups = nonZeroGroups.Where(c => c.Count > 0).ToList();

            foreach (var c in nonZeroGroups.ToList())
            {
                PageCpuTime prim = c.OrderBy(x => x.timestamp).FirstOrDefault();
                if ((prim.timestamp - startTime).TotalSeconds > 1)
                    nonZeroGroups.Remove(c);
            } //remove groups that start later than 1 second from menu click

            return nonZeroGroups;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            var d = new DirectoryInfo(Path.Combine(System.Windows.Forms.Application.StartupPath, textBox3.Text));
            var loggerFiles = Directory.GetFiles(d.FullName, "*.logger", SearchOption.AllDirectories);
            var cpuDataFiles = loggerFiles.Where(c => c.Contains("pdlrt")).ToList();

            foreach (var file in cpuDataFiles)
            {
                GetCPUFileData(file, out FileInfo dataFile, out List<float> percentagesList, out List<string> pagesList, out List<DateTime> dateTimesList, out List<PageCpuTime> pageCpuTimes);

                var terminus = "";
                if (dataFile.Name != "pdlrt.logger")
                    terminus = dataFile.Name.Replace(dataFile.Extension, "").Substring(6);
                else
                    terminus = "current";

                var oldtimerFiles = Directory.GetFiles(dataFile.DirectoryName, "timerData*", SearchOption.TopDirectoryOnly);
                foreach (var c in oldtimerFiles)
                {
                    var oldTimerFile = new FileInfo(c);
                    oldTimerFile.Delete();
                }

                ProcessGatheredCpuUsageData(percentagesList, pagesList, dateTimesList, terminus, dataFile.DirectoryName);
            }
        }

        private static void GetCPUFileData(string file, out FileInfo dataFile, out List<float> perc, out List<string> atPagesList, out List<DateTime> datetimes, out List<PageCpuTime> pageCpuTimes)
        {
            dataFile = new FileInfo(file);
            var textData = File.ReadAllLines(dataFile.FullName);
            perc = new List<float>();
            atPagesList = new List<string>();
            datetimes = new List<DateTime>();
            pageCpuTimes = new List<PageCpuTime>();
            for (int i = 0; i < textData.Length; i++)
            {
                var data = textData[i].Split(Convert.ToChar(" "))[3];
                var line = data.Split(Convert.ToChar(","));
                datetimes.Add(DateTime.Parse(line[0]));
                perc.Add(float.Parse(line[1]));
                atPagesList.Add(line[2]);
                pageCpuTimes.Add(new PageCpuTime()
                {
                    cpu = float.Parse(line[1]),
                    page = line[2],
                    timestamp = DateTime.Parse(line[0])
                });
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            NavigateHMIMenu();
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            Debug.Print(e.Node.Text);
            foreach (TreeNode node in e.Node.Nodes)
            {
                if (e.Node.Checked)
                    node.Checked = true;
                else
                    node.Checked = false;
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (collapsed)
            {
                treeView1.ExpandAll();
                collapsed = false;
            }
            else
            {
                treeView1.CollapseAll();
                collapsed = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            SelectOrDeselectAllTree();
        }

        private void SelectOrDeselectAllTree(bool useCheckbox = true)
        {
            bool myBool;
            if (useCheckbox)
                myBool = checkBox2.Checked;
            else
                myBool = true;

            foreach (TreeNode node in treeView1.Nodes)
            {
                node.Checked = myBool;
                foreach (TreeNode node2 in node.Nodes)
                {
                    node2.Checked = myBool;
                    foreach (TreeNode node3 in node2.Nodes)
                    {
                        node3.Checked = myBool;
                        foreach (TreeNode node4 in node3.Nodes)
                        {
                            node4.Checked = myBool;
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var d = new DirectoryInfo(Path.Combine(System.Windows.Forms.Application.StartupPath, textBox3.Text));
            if (!d.Exists)
            {
                listBox1.Items.Add("The folder " + d.Name + " was not found!");
                return;
            }

            var textFiles = Directory.GetFiles(d.FullName, "*.logger", SearchOption.AllDirectories);
            var timerFiles = textFiles.Where(c => c.Contains("timer")).ToList();

            List<List<PageTime>> alldata;
            List<string> fileList;
            GetGatheredDatasets(d, out alldata, out fileList);

            List<PageTime> computed = removeOutliersForEachPage(alldata, fileList);

            //List<PageTime> list = ComputeTimesFromDatasets(alldata, fileList);

            listView1.Columns.Add("Pdl", 230);
            listView1.Columns.Add("Loading Time [ms]", 110);
            listView1.Columns.Add("Min Time [ms]", 110);
            listView1.Columns.Add("Max Time [ms]", 110);
            listView1.View = View.Details;
            foreach (var c in computed)
            {
                //LogToFile(c.page + ", " + c.load + " ms", "\\stdDevTimerData.logger");
                string[] row = { c.page, Math.Round(c.load, 2).ToString(), c.min.ToString(), c.max.ToString() };
                var item1 = new ListViewItem(row);
                item1.Text = c.page;
                listView1.Items.Add(item1);
            }
            listView1.Refresh();
            checkBox3.Checked = true;
        }

        private List<PageTime> removeOutliersForEachPage(List<List<PageTime>> alldata, List<string> fileList)
        {
            List<PageTime> stdDevs = new List<PageTime>();
            foreach (var page in fileList)
            {
                var loadingTimes = new List<double>();
                foreach (var dataset in alldata)
                {

                    PageTime pageTime = dataset.FirstOrDefault(c => c.page == page);
                    if (pageTime != null)
                    {
                        loadingTimes.Add(pageTime.load);
                        LogToFile(page + "," + pageTime.load, "\\test.logger");
                    }
                }

                if (loadingTimes.Count > 0)
                {

                    var stdDev = StdDev(loadingTimes);

                    double average = loadingTimes.Average();
                    var someDoubles = loadingTimes.Where(c => c > average - stdDev && c < average + stdDev).OrderBy(c => c).ToList();
                    if (someDoubles.Count > 0)
                    {
                        var selectiveAverage = someDoubles.Average();

                        stdDevs.Add(new PageTime()
                        {
                            page = page,
                            load = selectiveAverage,
                            min = loadingTimes.Min(),
                            max = loadingTimes.Max(),
                        });
                    }
                }
            }

            return stdDevs;
        }

        private static void GetGatheredDatasets(DirectoryInfo d, out List<List<PageTime>> alldata, out List<string> fileList)
        {
            alldata = new List<List<PageTime>>();
            fileList = new List<string>();
            var textFiles = Directory.GetFiles(d.FullName, "*.logger", SearchOption.AllDirectories);
            var timerFiles = textFiles.Where(c => c.Contains("timer")).ToList();

            foreach (var file in timerFiles)
            {
                AddData(alldata, File.ReadAllLines(file));
            }

            foreach (var dir in d.GetDirectories())
            {
                var files = dir.GetFiles();
                var timerFile = files.FirstOrDefault(c => c.Name.StartsWith("timerData"));
                if (timerFile != null)
                    AddData(alldata, File.ReadAllLines(timerFile.FullName));
            }
            fileList = alldata.OrderByDescending(c => c.Count).FirstOrDefault().Select(c => c.page).Distinct().ToList();
        }

        private double StdDev(List<double> values)
        {
            double ret = 0;
            int count = values.Count();

            // i suppose i must get a std dev for each page in the list of values, therefore, the values list must be a list of lists
            if (count > 1)
            {
                //Compute the Average
                double avg = values.Average();

                //Perform the Sum of (value-avg)^2
                double sum = values.Sum(d => (d - avg) * (d - avg));

                //Put it all together
                ret = Math.Sqrt(sum / count);
            }
            return ret;
        }

        private static void AddData(List<List<PageTime>> alldata, string[] rawdata)
        {
            var data = new List<PageTime>();
            foreach (var line in rawdata)
            {
                var myline = line.Split(Convert.ToChar(","));
                var pageTime = new PageTime()
                {
                    page = myline[0].Split(Convert.ToChar("'"))[1],
                    load = double.Parse(myline[1].Replace(" ms", ""))
                };
                data.Add(pageTime);
            }

            alldata.Add(data);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Visible = checkBox3.Checked;

            if (!checkBox3.Checked)
            {
                Size = new Size()
                {
                    Height = 403,
                    Width = 406
                };
            }
            else
            {
                Size = new Size()
                {
                    Height = 403,
                    Width = 791
                };
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string path = System.Windows.Forms.Application.StartupPath + "\\ProcessedTimesExport.csv";
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            var items = listView1.Items;
            LogToFile("Page name,Measured loading time [ms]", "\\ProcessedTimesExport.csv", false);
            foreach (ListViewItem item in items)
            {
                LogToFile(item.SubItems[0].Text + "," + item.SubItems[1].Text + "," + item.SubItems[2].Text + "," + item.SubItems[3].Text, "\\ProcessedTimesExport.csv", false);
            }
        }

        private void Form1_DoubleClicked(object sender, EventArgs e)
        {
            if (Size.Width != 112)
            {
                Size = new Size()
                {
                    Width = 112,
                    Height = 107
                };
            }
            else
            {
                if (checkBox3.Checked)
                {
                    Size = new Size()
                    {
                        Width = 791,
                        Height = 363
                    };
                }
                else
                {
                    Size = new Size()
                    {
                        Width = 406,
                        Height = 363
                    };
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            InitTreeView();
        }

        #region comment
        //private void button6_Click_1(object sender, EventArgs e)
        //{
        //    var excel = new Microsoft.Office.Interop.Excel.Application();

        //    string fileName = @"C:\Users\MURA02\source\repos\WindowsFormsApp1\WinCC Timer\bin\Debug\Book1.xlsm";

        //    var d = new DirectoryInfo(@"C:\Users\MURA02\source\repos\WindowsFormsApp1\WinCC Timer\bin\Debug\SPM-1-Measurements");
        //    var textFiles = Directory.GetFiles(d.FullName, "*.logger", SearchOption.AllDirectories);
        //    var cpuFiles = textFiles.Where(c => c.Contains("pdlrt")).ToList();

        //    Microsoft.Office.Interop.Excel.Application oXL;
        //    Microsoft.Office.Interop.Excel._Workbook oWB;
        //    Microsoft.Office.Interop.Excel._Worksheet oSheet;
        //    Microsoft.Office.Interop.Excel._Worksheet oWS;

        //    try
        //    {


        //        //Start Excel and get Application object.
        //        oXL = new Microsoft.Office.Interop.Excel.Application();
        //        oXL.Visible = true;

        //        //Get a new workbook.
        //        oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
        //        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
        //        int i = 0;

        //        var allCpuData = new List<List<float>>();
        //        var allPageData = new List<List<string>>();
        //        var allDateTimeData = new List<List<DateTime>>();
        //        var allPageCpuTimes = new List<List<PageCpuTime>>();
        //        foreach (var file in cpuFiles)
        //        {
        //            GetCPUFileData(file, out FileInfo dataFile, out List<float> percentagesList, out List<string> pagesList, out List<DateTime> dateTimesList, out List<PageCpuTime> pageCpuTimes);
        //            allCpuData.Add(percentagesList);
        //            allPageData.Add(pagesList);
        //            allDateTimeData.Add(dateTimesList);
        //            allPageCpuTimes.Add(pageCpuTimes);

        //            i++;
        //            oSheet.get_Range("A" + i).Value2 = "Page";
        //            oSheet.get_Range("B" + i).get_Resize(1, pagesList.Count()).Value2 = pagesList.ToArray();
        //            i++;
        //            oSheet.get_Range("A" + i).Value2 = "Time";
        //            oSheet.get_Range("B" + i).get_Resize(1, dateTimesList.Count()).Value2 = dateTimesList.ToArray();
        //            oSheet.get_Range(i + ":" + i).NumberFormat = "hh:mm:ss.000";
        //            i++;
        //            oSheet.get_Range("A" + i).Value2 = "CPU";
        //            oSheet.get_Range("B" + i).get_Resize(1, percentagesList.Count()).Value2 = percentagesList.ToArray();
        //        }

        //        var distinctPages = new List<string>();
        //        foreach (var list in allPageData)
        //        {
        //            distinctPages.AddRange(list);
        //        }
        //        distinctPages = distinctPages.Distinct().Where(c => c != "").ToList();

        //        var pageData = new List<List<PageCpuTime>>();
        //        var sortedDataByPage = new List<List<List<PageCpuTime>>>();
        //        foreach (var page in distinctPages)
        //        {
        //            pageData = (from data in allPageCpuTimes
        //                        select data.Where(c => c.page == page).ToList()).ToList();
        //            sortedDataByPage.Add(pageData);
        //        }


        //        foreach (var c in sortedDataByPage[0])
        //        {
        //            //i++;
        //            //oSheet.get_Range("A" + i).Value2 = "Page";
        //            //oSheet.get_Range("B" + i).get_Resize(1, pagesList.Count()).Value2 = pagesList.ToArray();
        //            //i++;
        //            //oSheet.get_Range("A" + i).Value2 = "Time";
        //            //oSheet.get_Range("B" + i).get_Resize(1, dateTimesList.Count()).Value2 = dateTimesList.ToArray();
        //            //oSheet.get_Range(i + ":" + i).NumberFormat = "hh:mm:ss.000";
        //            //i++;
        //            //oSheet.get_Range("A" + i).Value2 = "CPU";
        //            //oSheet.get_Range("B" + i).get_Resize(1, percentagesList.Count()).Value2 = percentagesList.ToArray();
        //        }

        //        oXL.Visible = true;
        //        oXL.UserControl = true;
        //    }
        //    catch (Exception theException)
        //    {
        //        String errorMessage;
        //        errorMessage = "Error: ";
        //        errorMessage = String.Concat(errorMessage, theException.Message);
        //        errorMessage = String.Concat(errorMessage, " Line: ");
        //        errorMessage = String.Concat(errorMessage, theException.Source);

        //        MessageBox.Show(errorMessage, "Error");
        //    }
        //}

        //private void DisplayQuarterlySales(Microsoft.Office.Interop.Excel._Worksheet oWS)
        //{
        //    Microsoft.Office.Interop.Excel._Workbook oWB;
        //    Microsoft.Office.Interop.Excel.Series oSeries;
        //    Microsoft.Office.Interop.Excel.Range oResizeRange;
        //    Microsoft.Office.Interop.Excel._Chart oChart;
        //    String sMsg;
        //    int iNumQtrs;

        //    //Determine how many quarters to display data for.
        //    for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
        //    {
        //        sMsg = "Enter sales data for ";
        //        sMsg = String.Concat(sMsg, iNumQtrs);
        //        sMsg = String.Concat(sMsg, " quarter(s)?");

        //        DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
        //        MessageBoxButtons.YesNo);
        //        if (iRet == DialogResult.Yes)
        //            break;
        //    }

        //    sMsg = "Displaying data for ";
        //    sMsg = String.Concat(sMsg, iNumQtrs);
        //    sMsg = String.Concat(sMsg, " quarter(s).");

        //    MessageBox.Show(sMsg, "Quarterly Sales");

        //    //Starting at E1, fill headers for the number of columns selected.
        //    oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
        //    oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

        //    //Change the Orientation and WrapText properties for the headers.
        //    oResizeRange.Orientation = 38;
        //    oResizeRange.WrapText = true;

        //    //Fill the interior color of the headers.
        //    oResizeRange.Interior.ColorIndex = 36;

        //    //Fill the columns with a formula and apply a number format.
        //    oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
        //    oResizeRange.Formula = "=RAND()*100";
        //    oResizeRange.NumberFormat = "$0.00";

        //    //Apply borders to the Sales data and headers.
        //    oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
        //    oResizeRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

        //    //Add a Totals formula for the sales data and apply a border.
        //    oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
        //    oResizeRange.Formula = "=SUM(E2:E6)";
        //    oResizeRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle
        //    = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
        //    oResizeRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight
        //    = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

        //    CreateChart(oWS, out oWB, out oSeries, out oResizeRange, out oChart, iNumQtrs);

        //}

        //private static void CreateChart(Microsoft.Office.Interop.Excel._Worksheet oWS, out Microsoft.Office.Interop.Excel._Workbook oWB, out Microsoft.Office.Interop.Excel.Series oSeries, out Microsoft.Office.Interop.Excel.Range oResizeRange, out Microsoft.Office.Interop.Excel._Chart oChart, int iNumQtrs = 0)
        //{
        //    //Add a Chart for the selected data.
        //    oWB = (Microsoft.Office.Interop.Excel._Workbook)oWS.Parent;
        //    oChart = (Microsoft.Office.Interop.Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
        //    Missing.Value, Missing.Value);

        //    //Use the ChartWizard to create a new chart from the selected data.
        //    oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
        //    Missing.Value, iNumQtrs);
        //    oChart.ChartWizard(oResizeRange, Microsoft.Office.Interop.Excel.XlChartType.xl3DColumn, Missing.Value,
        //    Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
        //    Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //    oSeries = (Microsoft.Office.Interop.Excel.Series)oChart.SeriesCollection(1);
        //    oSeries.XValues = oWS.get_Range("A2", "A6");
        //    for (int iRet = 1; iRet <= iNumQtrs; iRet++)
        //    {
        //        oSeries = (Microsoft.Office.Interop.Excel.Series)oChart.SeriesCollection(iRet);
        //        String seriesName;
        //        seriesName = "=\"Q";
        //        seriesName = String.Concat(seriesName, iRet);
        //        seriesName = String.Concat(seriesName, "\"");
        //        oSeries.Name = seriesName;
        //    }

        //    oChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

        //    //Move the chart so as not to cover your data.
        //    oResizeRange = (Microsoft.Office.Interop.Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
        //    oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
        //    oResizeRange = (Microsoft.Office.Interop.Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
        //    oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        //}
        #endregion

        private void numberOfMeasurements_TextChanged(object sender, EventArgs e)
        {
            GetMenuData();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            RunTasksForScreenshots();
        }

        private void ScourPage()
        {
            IntPtr handle = GetHandle("PDLRTisAliveAndWaitsForYou", "WinCC-Runtime - ");
            if (handle == IntPtr.Zero)
            {
                handle = WndSearcher.SearchForWindow("GRAFClass", "Graphics Designer - "); //partial search for window
                if (handle == IntPtr.Zero)
                    return;
            }

            //Bitmap img = TheMagic.GetPngByHandle(handle);

            PInvokeLibrary.SetForegroundWindow(handle);
            string mainScreen = GetMainScreen();
            currentPage = mainScreen;
            ScreenshotAndSave(true);

            var extractedData = new List<ObjectData>();

            if (popupsCheckbox.Checked)
            {
                GetRuntimeInfo("HMIButton", out extractedData, "@");

                FindOpenClosePopups(handle, extractedData.Where(c => !c.CallsPdl.Contains("_e_")).ToList());
            }

            if (dropCheckbox.Checked)
            {
                if (extractedData.Count == 0)
                {
                    GetRuntimeInfo("HMIButton", out extractedData, "FlyoutButton");
                }

                FindOpenCloseFlyouts(handle, extractedData.Where(c => c.ObjectName.Contains("FlyoutButton") && !c.CallsPdl.Contains("_e_")).ToList(), false);
            }

            if (embeddedsBox.Checked)
            {
                List<ObjectData> extractedData2 = new List<ObjectData>();
                if (extractedData.Count == 0)
                {
                    GetRuntimeInfo("HMIButton", out extractedData);
                }
                GetRuntimeInfo("HMIGroup", out extractedData2);

                extractedData = extractedData.Where(c => c.CallsPdl.Contains("_e_")).ToList();
                extractedData2 = extractedData2.Where(c => c.CallsPdl.Contains("_e_")).ToList();

                extractedData = extractedData.Concat(extractedData2).ToList();

                FindSwitchEmbeddeds(handle, extractedData); //embeddeds will change pages; these also need to be checked for popups and embeddeds 
            }
        }
        private List<string> navigatedPages = new List<string>();



        private string GetMainScreen()
        {
            var list = new List<string>();
            GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen);
            foreach (IHMIScreen s in screens)
            {
                list.Add(s.ObjectName);
            }
            var currentNormal = list?.FirstOrDefault(c => c.ToUpper().Contains("_n_".ToUpper()));
            var currentWide = list?.FirstOrDefault(c => c.ToUpper().Contains("_w_".ToUpper()));

            var mainScreen = currentNormal != null ? currentNormal : currentWide;
            return mainScreen;
        }

        private void FindOpenCloseFlyouts(IntPtr handle, List<ObjectData> extractedData, bool first = true)
        {
            foreach (ObjectData p in extractedData)
            {
                int midPointX = (p.RealLeft + p.RealRight) / 2;
                int midPointY = (p.RealTop + p.RealBottom) / 2;

                RobustClick(handle, midPointX, midPointY, 1); Thread.Sleep(3000);
                navigatedPages.Add(p.CallsPdl);

                GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen);
                var s = screens.Cast<IHMIScreen>();
                IHMIScreen activePopup = s.FirstOrDefault(c => c.ObjectName.Contains("_p_"));

                if (activePopup != null)
                {
                    bool isMyPage = p.CallsPdl.Contains(activePopup?.ObjectName);
                    //take screenshot here
                    SetDateString();
                    currentPage = activePopup.ObjectName;
                    PInvokeLibrary.SetForegroundWindow(handle);


                    ScreenshotAndSave(true,
                        new Rectangle()
                        {
                            Width = activePopup.Parent.Width,
                            Height = activePopup.Parent.Height,
                            X = activePopup.Parent.Left,
                            Y = activePopup.Parent.Top
                        });

                    RobustClick(handle, midPointX, midPointY, 1);
                    Thread.Sleep(3000);
                }
            }
        }

        private void FindOpenClosePopups(IntPtr handle, List<ObjectData> extractedData)
        {
            var openedPdls = new List<string>();

            var relevantData = extractedData.Where(c => c.ObjectName.StartsWith("@V3_SMS") &&
            !c.CallsPdl.StartsWith("SYS#") &&
            !c.CallsPdl.StartsWith("MED#") &&
            !c.CallsPdl.StartsWith("@sms"));

            foreach (var p in relevantData)
            {
                int midPointX = (p.RealLeft + p.RealRight) / 2;
                int midPointY = (p.RealTop + p.RealBottom) / 2;

                RobustClick(handle, midPointX, midPointY, 1); Thread.Sleep(3000);
                navigatedPages.Add(p.CallsPdl);

                GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen);
                var s = screens.Cast<IHMIScreen>();
                IHMIScreen activePopup = s.FirstOrDefault(c => c.ObjectName.Contains("_p_"));

                if (activePopup != null)
                {
                    bool isMyPage = p.CallsPdl.Contains(activePopup?.ObjectName);

                    openedPdls.Add(activePopup?.ObjectName);

                    //take screenshot here
                    currentPage = activePopup.ObjectName;
                    PInvokeLibrary.SetForegroundWindow(handle);

                    ScreenshotAndSave(
                        true,
                        new Rectangle()
                        {
                            Width = activePopup.Parent.Width,
                            Height = activePopup.Parent.Height,
                            X = activePopup.Parent.Left,
                            Y = activePopup.Parent.Top
                        });

                    ClosePopup(handle);

                    currentActiveScreen = ""; Thread.Sleep(1000);
                }
                else
                {
                    Console.WriteLine("Failed to open popup " + p.Page);
                }
            }
        }

        private static void ClosePopup(IntPtr handle)
        {
            Bitmap img = TheMagic.GetPngByHandle(handle);
            List<TheMagic.PosBitmap> closes = TheMagic.Find(img, TheMagic.MakeExistingTransparent((Bitmap)Resources.ResourceManager.GetObject("close")), "close");
            foreach (var c in closes)
            {
                int midPointX = c.Left + c.Width / 2;
                int midPointY = c.Top + c.Height / 2;
                ClickAtXY(midPointX, midPointY); Thread.Sleep(1000);
            }

            //return img;
        }

        private void GetRuntimeInfo(string type, out List<ObjectData> dataList, string nameContains = "")
        {
            var objects = new List<grafexe.HMIObject>();
            dataList = new List<ObjectData>();

            GetScreenInfo(out List<IHMIScreen> filteredScreens, out List<WindowData> windowData);

            dataList = GetObjectsInfo(type, nameContains, filteredScreens, windowData);
        }

        private List<ObjectData> GetObjectsInfo(string type, string nameContains, List<IHMIScreen> filteredScreens, List<WindowData> windowData)
        {
            var g = new grafexe.Application();
            var dataList = new List<ObjectData>();

            foreach (var s in filteredScreens)
            {
                ////this used to yield error. check for alternative to find picture window position in the frame
                //try
                //{
                    //LogToFile(s.ObjectName + " at " + s.Parent.Left + ", " + s.Parent.Top + ", " + s.Parent.Width + ", " + s.Parent.Height, "\\Screen.log", false);
                //}
                //catch (Exception ex)
                //{
                //    LogToFile(ex.Message, "\\Screen.log", false);
                //}

                var mainScreen = GetMainScreen();
                WindowData screenData = new WindowData();

                var mainScreenData = windowData.FirstOrDefault(c => c.Parent == mainScreen);
                if (mainScreenData == null)
                {
                    screenData = windowData.FirstOrDefault(c => c.Name == s.Parent.ObjectName);
                }
                else
                {
                    screenData = windowData.FirstOrDefault(c => c.Parent == mainScreen);
                }
                screenData = windowData.FirstOrDefault(c => c.Name == s.Parent.ObjectName);

                
                var left = /*s.Parent.Parent.Parent.Width == 1920 && !s.Parent.Parent.ObjectName.StartsWith("@") ? screenData.ScreenLeft - 175 :*/ screenData.ScreenLeft;
                //if (left != screenData.ScreenLeft && s.Parent.Parent.Parent.Width == 1920 && !s.Parent.Parent.ObjectName.StartsWith("@"))
                //{
                //    screenData.ScreenLeft = left;
                //}
                var top = screenData.ScreenTop;

                var screenItems = s.ScreenItems.Cast<IHMIScreenItem>();
                var query = screenItems.Where(o => o.ObjectName.Contains(nameContains) && o.Type == type).ToList();

                foreach (IHMIScreenItem o in query) //screenitems is rt objects
                {
                    var obj = RetrieveObjectProperties(s.ObjectName, o.ObjectName, g);

                    if (obj != null)
                    {
                        listBox1.Items.Add(obj.ObjectName.value + "," + obj.Left.value + "," + obj.Top.value);

                        //LogToFile(o.Parent.ObjectName + "," + obj.ObjectName.value + "," + obj.Left.value + "," + obj.Top.value, "\\Screen.log", false);
                        //LogToFile(o.Parent.ObjectName + "," + obj.ObjectName.value + "," +
                        //    (left + obj.Left.value) + "," +
                        //    (top + obj.Top.value) + "," +
                        //    (left + obj.Left.value + obj.Width.value) + "," +
                        //    (top + obj.Top.value + obj.Height.value), "\\Screen.log", false);


                        var events = obj.Events.Cast<HMIEvent>();
                        HMIEvent evs = events.FirstOrDefault(e => e.EventName == "OnLButtonUp" && e.Actions.Count > 0);

                        string pdl = "";

                        if (evs == null)
                        {
                            //break;
                        }
                        else
                        {

                            HMIActions acts = evs.Actions;

                            foreach (HMIScriptInfo act in acts)
                            {
                                if (pdl == "")
                                {
                                    var vb = act.SourceCode;

                                    var callingRow = vb.Split("\r\n".ToCharArray()).FirstOrDefault(c => c.Contains("pdlName = "));
                                    pdl = callingRow?.Split("\"".ToCharArray())?.ElementAtOrDefault(1);

                                    var embeddedRow = vb.Split("\r\n".ToCharArray()).FirstOrDefault(c => c.Contains(".PictureName = "));
                                    if (embeddedRow != null)
                                    {
                                        pdl = embeddedRow?.Split("\"".ToCharArray())?.ElementAtOrDefault(1);
                                    }
                                    if (pdl != "")
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        pdl = pdl != null ? pdl.Replace(".pdl", "") : "";

                        //dataList.FirstOrDefault(c => c.ObjectName == o.ObjectName).CallsPdl = pdl != "" ? pdl : "";

                        if (pdl != "")
                        {
                            dataList.Add(new ObjectData()
                            {
                                RealLeft = left + obj.Left.value,
                                RealRight = left + obj.Left.value + obj.Width.value,
                                RealTop = top + obj.Top.value,
                                RealBottom = top + obj.Top.value + obj.Height.value,

                                OffsetLeft = left,
                                OffsetTop = top,
                                ObjectName = obj.ObjectName.value,
                                Page = s.ObjectName,
                                CallsPdl = pdl
                            });
                            Console.WriteLine(pdl);
                        }

                        //LogToFile(obj.ObjectName.value + " calls " + dataList.Last().CallsPdl, "\\PdlCalls.log", false);
                    }
                }
            }

            dataList = dataList.Where(c => c.CallsPdl != "" && c.CallsPdl != null && !navigatedPages.Contains(c.CallsPdl)).ToList();
            return dataList;
        }

        private void GetScreenInfo(out List<IHMIScreen> relevantScreens, out List<WindowData> windowData)
        {
            relevantScreens = new List<IHMIScreen>();

            GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen screen);

            List<IHMIScreen> screensList = screens.Cast<IHMIScreen>().ToList();

            IEnumerable<IHMIScreenItem> parents = screensList.Where(c => c.Parent != null).Select(c => c.Parent);
            IEnumerable<IHMIScreenItem> menus = screensList.Where(c => c.Parent != null && c.Parent.ObjectName.Contains("menu")).Select(c => c.Parent);
            IEnumerable<string> accessPaths = screensList.Where(c => c.Parent != null).Select(c => c.AccessPath);

            string mainscreen = GetMainScreen();

            var allScreenNames = screensList.Select(c => c.ObjectName).ToList();

            //relevantScreens = screensList.Where(c => c.Parent != null && (c.Parent.ObjectName.StartsWith("@DataClass") || c.Parent.ObjectName == "@ProcAreaLight"));
            foreach (IHMIScreen s in screensList)
            {
                //if (s.Parent != null)
                //{
                //    if (s.Parent.ObjectName.StartsWith("@DataClass") || s.Parent.Type == "HMIScreenWindow")
                //    {
                //        //if (s.Parent.Parent.ObjectName == "@frame_s_content")
                //        //{
                //        relevantScreens.Add(s);
                //        //}
                //    }
                //    //else if (s.Parent.ObjectName == "@DataClass3" && s.Parent.Parent.ObjectName == mainscreen)
                //    //{
                //    //    relevantScreens.Add(s);
                //    //}

                //    if (s.Parent.ObjectName == "@ProcAreaLight")
                //    {
                //        relevantScreens.Add(s);
                //    }
                //}

                var hashtagSplit = s.ObjectName.Split("#".ToCharArray());

                if (hashtagSplit.Length > 1)
                {
                    var dashSplit = hashtagSplit[1].Split("-".ToCharArray());

                    if (dashSplit.Length > 1)
                    {
                        var parseSuccess = int.TryParse(dashSplit[0], out int parseResult);

                        if (parseSuccess == true)
                        {
                            Console.WriteLine(parseResult);
                            if (parseResult > 0 && !s.ObjectName.ToUpper().Contains("_f_".ToUpper()))
                            {
                                relevantScreens.Add(s);
                            }
                        }

                    }
                }

            }

            var result = from c in relevantScreens
                         join c2 in relevantScreens on c.Parent.ObjectName equals c2.Parent.ObjectName
                         where c != c2
                         select c;
            if (result.Count() > 0)
            {
                Console.WriteLine("Found some duplicates at c.parent.objectname boss");
            }

            var screenNames = relevantScreens.Select(c => c.ObjectName).ToList();

            var screenWindowsQuery = relevantScreens.Select(c => c.Parent).ToList();
            windowData = GetWindowData(screenWindowsQuery);
        }

        private static List<WindowData> GetWindowData(List<IHMIScreenItem> screenWindowsQuery)
        {
            List<WindowData> windowData = new List<WindowData>();

            var t = screenWindowsQuery.Where(c => c.Parent != null).Where(c => c.Parent.Parent != null).Select(c => c.Parent.Parent).ToList();

            foreach (IHMIScreenItem s in screenWindowsQuery)
            {
                int recLeft = 0;
                int recTop = 0;
                GetPositionRec(ref recLeft, ref recTop, s);

                windowData.Add(new WindowData()
                {
                    Width = s.Width,
                    Top = s.Top,
                    Height = s.Height,
                    Left = s.Left,
                    Name = s.ObjectName,
                    Visible = s.Visible,
                    Type = s.Type,
                    Layer = s.Layer,
                    Parent = s.Parent.ObjectName,
                    ScreenLeft = recLeft,
                    ScreenTop = recTop,
                    Enabled = s.Enabled
                });
            }

            var dc1 = windowData.FirstOrDefault(c => c.Name == "@DataClass1");
            var dc2 = windowData.FirstOrDefault(c => c.Name == "@DataClass2");
            var dc3 = windowData.FirstOrDefault(c => c.Name == "@DataClass3");

            //if (dc1 != null && dc2 != null)
            //{
            //    if (dc2.Width == 1920)
            //    {
            //        dc2.Left -= dc1.Width;
            //        dc2.ScreenLeft -= dc1.Width;
            //        if (dc3 != null)
            //        {
            //            dc3.Left -= dc1.Width;
            //            dc3.ScreenLeft -= dc1.Width;
            //        }
            //    }
            //}

            return windowData;
        }

        private static void GetPositionRec(ref int recLeft, ref int recTop, IHMIScreenItem screenWindow)
        {
            if (screenWindow != null)
            {
                recLeft += screenWindow.Left;
                recTop += screenWindow.Top;
                if (screenWindow.Parent.Parent != null)
                {
                    if (screenWindow.ObjectName == screenWindow.Parent.Parent.ObjectName)
                    {
                        recLeft -= 175;
                    }
                    GetPositionRec(ref recLeft, ref recTop, screenWindow.Parent.Parent);
                }
            }
        }

        private void FindSwitchEmbeddeds(IntPtr handle, List<ObjectData> extractedData, bool innerEmbeeddeds = false)
        {
            ObjectData o = extractedData.FirstOrDefault();

            if (o != null)
            {

                int midPointX = (o.RealLeft + o.RealRight) / 2;
                int midPointY = (o.RealTop + o.RealBottom) / 2;

                RobustClick(handle, midPointX, midPointY, 1); Thread.Sleep(2000);
                navigatedPages.Add(o.CallsPdl);

                var after = GetCurrentRuntimeFilelist();
                //var after = RuntimeFunctions.GetCurrentRuntimeFilelist();

                var emb = after.FirstOrDefault(c => c.ObjectName == o.CallsPdl);

                if (emb == null)
                {
                    Console.WriteLine("yep");
                }
                else
                {
                    var filteredScreens = new List<IHMIScreen>();
                    filteredScreens.Add(emb);

                    var windowData = GetWindowData(filteredScreens?.Where(c => c.Parent != null)?.Select(c => c.Parent)?.ToList());
                    var dataList = GetObjectsInfo("HMIButton", "", filteredScreens, windowData);
                    var dataList2 = GetObjectsInfo("HMIGroup", "", filteredScreens, windowData);
                    dataList = dataList.Concat(dataList2).ToList();

                    dataList = dataList.Where(c => c.CallsPdl != "" && c.CallsPdl != null && !navigatedPages.Contains(c.CallsPdl))
                                       .Where(c => c.CallsPdl.ToUpper().Contains("_e_".ToUpper()))
                                       .ToList();

                    ReadActiveScreen();

                    //take screenshot here
                    SetDateString();
                    currentPage = o.CallsPdl;

                    var embdata = windowData.FirstOrDefault();

                    PInvokeLibrary.SetForegroundWindow(handle);

                    ScreenshotAndSave(true);
                        //true,
                        //new Rectangle()
                        //{
                        //    Width = embdata.Width,
                        //    Height = embdata.Height,
                        //    X = embdata.ScreenLeft,
                        //    Y = embdata.ScreenTop
                        //});

                    if (dataList.Count > 0)
                    {
                        FindSwitchEmbeddeds(handle, dataList, true);
                    }
                }
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            ScourPage();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Q | Keys.Control))
            {
                Environment.Exit(0);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            GetCurrentRuntimeFilelist();
        }

        private List<IHMIScreen> GetCurrentRuntimeFilelist()
        {
            IHMIScreens screens;
            IHMIScreen activeScreen;
            GetRuntimeScreens(out screens, out activeScreen);

            var foundScreens = new List<IHMIScreen>();

            LogToFile(activeScreen.AccessPath, "\\Screens.txt", false);

            //listBox1.Items.Add(activeScreen.ObjectName);

            var rt = new CCHMIRUNTIME.HMIRuntime();

            var screensList = screens.Cast<IHMIScreen>();
            List<CCHMIRUNTIME.IHMIDataSet> datasets = screensList.
                Where(c => c.Parent != null).
                Where(c => c.Parent.ObjectName.StartsWith("@DataClass")).
                Select(c => c.DataSet).
                ToList();

            foreach (CCHMIRUNTIME.IHMIScreen s in screens)
            {
                LogToFile(s.ObjectName, "\\Screens.txt", false);

                foundScreens.Add(s);
            }

            return foundScreens;
        }

        private void GetRuntimeScreens(out IHMIScreens screens, out IHMIScreen activeScreen)
        {
            CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();
            CCHMIRTGRAPHICS.HMIRTGraphics graphics = new CCHMIRTGRAPHICS.HMIRTGraphics();

            screens = rt.Screens;
            activeScreen = rt.ActiveScreen;

            currentActiveScreen = activeScreen.ObjectName;
            currentPage = currentActiveScreen;
        }

        private string ReadActiveScreen()
        {
            CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();
            currentActiveScreen = rt.ActiveScreen.ObjectName;
            currentPage = currentActiveScreen;

            return currentActiveScreen;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add(GetMainScreen());
        }

        private grafexe.HMIObject RetrieveObjectProperties(string pdlName, string objName, grafexe.Application g, grafexe.HMIOpenDocumentType openType = grafexe.HMIOpenDocumentType.hmiOpenDocumentTypeInvisible)
        {
            string seldocfullname = Path.Combine(g.ApplicationDataPath, pdlName + ".pdl");

            listBox1.Items.Add("Opening... " + seldocfullname);

            grafexe.HMIObject go = null;

            try
            {
                grafexe.Document seldoc = g.Documents.Open(seldocfullname, openType);
                grafexe.HMIObjects selos = seldoc.HMIObjects;

                go = selos.Find(ObjectName: objName).Count > 0 ? selos.Find(ObjectName: objName)[1] : null;

                if (go != null)
                {
                    if (go.Visible.value == false && go.Visible.DynamicStateType == HMIDynamicStateType.hmiDynamicStateTypeNoDynamic)
                        go = null;
                }

                if (go != null)
                {
                    listBox1.Items.Add(go.ObjectName.value);
                }

                return go;
            }
            catch (Exception exc)
            {
                LogToFile(exc.Message, logName);
            }
            finally
            {
            }
            return null;
        }
    }
}

public class ObjectData
{
    public int OffsetLeft { get; internal set; }
    public int OffsetTop { get; internal set; }
    public int RealLeft { get; internal set; } //relative to whole screen
    public int RealTop { get; internal set; } //relative to whole screen
    public int RealRight { get; internal set; } //relative to whole screen
    public int RealBottom { get; internal set; } //relative to whole screen
    public string Page { get; internal set; }
    public string ObjectName { get; internal set; }
    public string CallsPdl { get; internal set; }
}

public class WindowData
{
    public int Top { get; internal set; }
    public int Left { get; internal set; }
    public int Width { get; internal set; }
    public int Height { get; internal set; }
    public string Name { get; internal set; }
    public bool Visible { get; internal set; }
    public string Type { get; internal set; }
    public int Layer { get; internal set; }
    public bool Enabled { get; internal set; }
    public string Parent { get; internal set; }
    public int ScreenLeft { get; internal set; }
    public int ScreenTop { get; internal set; }
}

public class PageTime
{
    public string page { get; internal set; }
    public double load { get; internal set; }
    public double min { get; internal set; }
    public double max { get; internal set; }
}

public class PageCpuTime
{
    public string page { get; internal set; }
    public float cpu { get; internal set; }
    public DateTime timestamp { get; internal set; }
}

public class MenuRow
{
    public string ID { get; internal set; }
    public string RefId { get; internal set; }
    public string Layer { get; internal set; }
    public string Pos { get; internal set; }
    public string LCID { get; internal set; }
    public string ParentId { get; internal set; }
    public string Caption { get; internal set; }
    public string Flags { get; internal set; }
    public string Pdl { get; internal set; }
    public string Parameter { get; internal set; }
}

public class MyCheckBox : CheckBox
{
    public MyCheckBox()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        Padding = new Padding(6);
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        this.OnPaintBackground(e);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using (var path = new GraphicsPath())
        {
            var d = Padding.All;
            var r = this.Height - 2 * d;
            path.AddArc(d, d, r, r, 60, 120);
            path.AddArc(this.Width - r - d, d, r, r, -60, 120);
            path.CloseFigure();
            e.Graphics.FillPath(Checked ? Brushes.White : Brushes.White, path);
            r = Height - 1;
            var rect = Checked ? new Rectangle(Width - r - 1, 0, r, r)
                               : new Rectangle(0, 0, r, r);
            e.Graphics.FillEllipse(Checked ? Brushes.Blue : Brushes.Gray, rect);
        }
    }
}