using WindowHandle;
using WindowsUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Interoperability;
using System.Management;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Management.Automation.Runspaces;
using System.Linq.Expressions;

namespace AutomateDownloader
{
    class NCMForm : Form
    {
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox firstClientIndexBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.TextBox numClTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox pathTextBox;
        private System.Windows.Forms.TextBox ipTextBox;
        private System.Windows.Forms.TextBox unTextBox;
        private System.Windows.Forms.TextBox passTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox rdpCheckBox;
        private System.Windows.Forms.GroupBox rdpBox1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.TextBox heightBox;
        private System.Windows.Forms.TextBox widthBox;
        private System.Windows.Forms.TextBox topBox;
        private System.Windows.Forms.TextBox leftBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private Button button2;
        private TextBox textBox1;
        private TextBox textBox2;
        private GroupBox groupBox1;
        private Form frm1 = new Form();

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new NCMForm());
            //Done by Muresan Radu-Adrian MURA02 20200716
        }

        private List<string> ipList = new List<string>();
        private List<string> sdList = new List<string>();

        public NCMForm()
        {
            CloseConflictingProcesses("NCM_Downloader");

            InitializeComponent();

            string configPath = string.Empty;
            if (File.Exists(Application.StartupPath + "\\NCM_Downloader.ini"))
            {
                configPath = Application.StartupPath + "\\NCM_Downloader.ini";
            }
            else if (File.Exists(Application.StartupPath + "\\NCM_Downloader.exe.ini"))
            {
                configPath = Application.StartupPath + "\\NCM_Downloader.exe.ini";
            }
            if (configPath != string.Empty)
            {
                GetInitialValues(configPath);
            }

            SetTooltips();

            RenewIpsOrInit();

            button1.Click += new EventHandler(Button1_Click);

            this.DoubleClick += new EventHandler(Form1_DoubleClick);

            this.Controls.Add(firstClientIndexBox);
        }

        private void GetInitialValues(string path)
        {
            var configFile = File.ReadLines(path);
            var fileLen = configFile.Count();

            pathTextBox.Text = configFile.ElementAt(0);
            if (fileLen >= 2) firstClientIndexBox.Text = configFile.ElementAt(1);
            if (fileLen >= 3) numClTextBox.Text = configFile.ElementAt(2);
            if (fileLen >= 4) ipTextBox.Text = configFile.ElementAt(3);
            if (fileLen >= 5) unTextBox.Text = configFile.ElementAt(4);
            if (fileLen >= 6) passTextBox.Text = configFile.ElementAt(5);
            if (fileLen >= 7) rdpCheckBox.Checked = Convert.ToBoolean(configFile.ElementAt(6));
            if (fileLen >= 8) widthBox.Text = configFile.ElementAt(7);
            if (fileLen >= 9) heightBox.Text = configFile.ElementAt(8);
            if (fileLen >= 10) leftBox.Text = configFile.ElementAt(9);
            if (fileLen >= 11) topBox.Text = configFile.ElementAt(10);
        }

        private void RenewIpsOrInit()
        {
            GetIpsLmHosts();
            if (ipList.Count() > 0)
            {
                checkedListBox1.Items.Clear();
                foreach (var item in ipList)
                {
                    checkedListBox1.Items.Add(item.Split(Convert.ToChar("\t"))[1]);
                }
                numClTextBox.Text = Convert.ToString(ipList.Count());
                firstClientIndexBox.Text = Convert.ToString(sdList.Count() + 1);
                statusLabel.Text = "Ready";
            }
            else
            {
                statusLabel.Text = "Status: Did not find LmHosts file, please check path";
            }
            statusLabel.Refresh();
            checkedListBox1.Refresh();
        }

        public string GetControlText(IntPtr hWnd)
        {

            // Get the size of the string required to hold the window title (including trailing null.) 
            Int32 titleSize = PInvokeLibrary.SendMessage((int)hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, 0).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

            return title.ToString();
        }

        #region TextExtractArea


        // privileges
        const int PROCESS_CREATE_THREAD = 0x0002;
        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int PROCESS_VM_OPERATION = 0x0008;
        const int PROCESS_VM_WRITE = 0x0020;
        const int PROCESS_VM_READ = 0x0010;

        // used for memory allocation
        const uint MEM_COMMIT = 0x00001000;
        const int MEM_DECOMMIT = 0x4000;
        const uint MEM_RESERVE = 0x00002000;
        const uint PAGE_READWRITE = 4;

        ///<summary>Returns the tree node information from another process.</summary>
        ///<param name="hwndItem">Handle to a tree node item.</param>
        ///<param name="hwndTreeView">Handle to a tree view control.</param>
        ///<param name="process">Process hosting the tree view control.</param>
        ///
        private void Testing(IntPtr ncmHandle, string anyPopupClass)
        {
            //testing for treeview
            var parent3 = PInvokeLibrary.FindWindowEx(ncmHandle, IntPtr.Zero, "MDIClient", null);
            var parent2 = PInvokeLibrary.FindWindowEx(parent3, IntPtr.Zero, "Afx:400000:b:10003:6:104d09c9", null);
            var parent1 = PInvokeLibrary.FindWindowEx(parent2, IntPtr.Zero, "AfxFrameOrView42", null);
            var parent = PInvokeLibrary.FindWindowEx(parent1, IntPtr.Zero, anyPopupClass, "");
            var treeH = PInvokeLibrary.FindWindowEx(parent, IntPtr.Zero, "SysTreeView32", "Generic1");
            var treeItemHeight = (int)PInvokeLibrary.SendMessage(treeH, (int)WindowsMessages.TVM_GETITEMW, 5, IntPtr.Zero); //works

            //PInvokeLibrary.TVITEM item = new PInvokeLibrary.TVITEM();

            var t = AllocTest(Process.GetProcessById(7396), treeH, IntPtr.Zero);

            //instead build an xml

            var s = GetTreeItemText(treeH, IntPtr.Zero);

            Console.WriteLine("uhum");
        }

        private static NodeData AllocTest(Process process, IntPtr hwndTreeView, IntPtr hwndItem)
        {
            // code based on article posted here: http://www.codingvision.net/miscellaneous/c-inject-a-dll-into-a-process-w-createremotethread

            // handle of the process with the required privileges
            IntPtr procHandle = PInvokeLibrary.OpenProcess((PInvokeLibrary.ProcessAccessFlags)(PROCESS_CREATE_THREAD | PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ), false, process.Id);

            // Write TVITEM to memory
            // Invoke TVM_GETITEM
            // Read TVITEM from memory

            var item = new TVITEMEX();
            item.hItem = hwndItem;
            item.mask = (int)(WindowsMessages.TVIF_HANDLE | /*WindowsMessages.TVIF_CHILDREN |*/ WindowsMessages.TVIF_TEXT);
            item.cchTextMax = 1024;
            item.pszText = PInvokeLibrary.VirtualAllocEx(procHandle, IntPtr.Zero, (uint)item.cchTextMax, (PInvokeLibrary.AllocationType)(MEM_COMMIT | MEM_RESERVE), (PInvokeLibrary.MemoryProtection)PAGE_READWRITE); // node text pointer

            byte[] data = getBytes(item);

            int dwSize = (int)data.Length;
            IntPtr allocMemAddress = PInvokeLibrary.VirtualAllocEx(procHandle, IntPtr.Zero, (uint)dwSize, (PInvokeLibrary.AllocationType)(MEM_COMMIT | MEM_RESERVE), (PInvokeLibrary.MemoryProtection)PAGE_READWRITE); // TVITEM pointer

            int nSize = dwSize;
            IntPtr bytesWritten;
            bool successWrite = PInvokeLibrary.WriteProcessMemory(procHandle, allocMemAddress, data, nSize, out bytesWritten);

            var sm = PInvokeLibrary.SendMessage(hwndTreeView, (int)WindowsMessages.TVM_GETITEM, 0, allocMemAddress);

            IntPtr bytesRead;
            bool successRead = PInvokeLibrary.ReadProcessMemory(procHandle, allocMemAddress, data, nSize, out bytesRead);

            IntPtr bytesReadText;
            byte[] nodeText = new byte[item.cchTextMax];
            bool successReadText = PInvokeLibrary.ReadProcessMemory(procHandle, item.pszText, nodeText, (int)item.cchTextMax, out bytesReadText);

            bool success1 = PInvokeLibrary.VirtualFreeEx(procHandle, allocMemAddress, dwSize, (PInvokeLibrary.AllocationType)MEM_DECOMMIT);
            bool success2 = PInvokeLibrary.VirtualFreeEx(procHandle, item.pszText, (int)item.cchTextMax, (PInvokeLibrary.AllocationType)MEM_DECOMMIT);

            var item2 = fromBytes<TVITEMEX>(data);

            String name = Encoding.Unicode.GetString(nodeText);
            int x = name.IndexOf('\0');
            if (x >= 0)
                name = name.Substring(0, x);

            NodeData node = new NodeData();
            node.Text = name;
            node.HasChildren = (item2.cChildren == 1);

            return node;
        }

        public class NodeData
        {
            public String Text { get; set; }
            public bool HasChildren { get; set; }
        }

        private static byte[] getBytes(Object item)
        {
            int size = Marshal.SizeOf(item);
            byte[] arr = new byte[size];
            IntPtr ptr = Marshal.AllocHGlobal(size);

            Marshal.StructureToPtr(item, ptr, true);
            Marshal.Copy(ptr, arr, 0, size);
            Marshal.FreeHGlobal(ptr);

            return arr;
        }

        private static T fromBytes<T>(byte[] arr)
        {
            T item = default(T);
            int size = Marshal.SizeOf(item);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            Marshal.Copy(arr, 0, ptr, size);
            item = (T)Marshal.PtrToStructure(ptr, typeof(T));
            Marshal.FreeHGlobal(ptr);
            return item;
        }

        // Note: different layouts required depending on OS versions.
        // https://msdn.microsoft.com/en-us/library/windows/desktop/bb773459%28v=vs.85%29.aspx
        [StructLayout(LayoutKind.Sequential)]
        public struct TVITEMEX
        {
            public uint mask;
            public IntPtr hItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;
            public int iIntegral;
            public uint uStateEx;
            public IntPtr hwnd;
            public int iExpandedImage;
            public int iReserved;
        }


        public static string GetTreeItemText(IntPtr treeViewHwnd, IntPtr hItem)
        {
            string itemText;

            PInvokeLibrary.GetWindowThreadProcessId(treeViewHwnd, out uint pid);

            IntPtr process = PInvokeLibrary.OpenProcess(PInvokeLibrary.ProcessAccessFlags.VirtualMemoryOperation | PInvokeLibrary.ProcessAccessFlags.VirtualMemoryRead | PInvokeLibrary.ProcessAccessFlags.VirtualMemoryWrite | PInvokeLibrary.ProcessAccessFlags.QueryInformation, false, (int)pid);
            if (process == IntPtr.Zero)
                throw new Exception("Could not open handle to owning process of TreeView", new Win32Exception());

            try
            {
                uint tviSize = (uint)Marshal.SizeOf(typeof(PInvokeLibrary.TVITEM));

                uint textSize = (uint)WindowsMessages.MY_MAXLVITEMTEXT;
                bool isUnicode = PInvokeLibrary.IsWindowUnicode(treeViewHwnd);
                if (isUnicode)
                    textSize *= 2;

                IntPtr tviPtr = IntPtr.Zero;
                try
                {
                    tviPtr = PInvokeLibrary.VirtualAllocEx(process, IntPtr.Zero, tviSize + textSize, PInvokeLibrary.AllocationType.Commit, PInvokeLibrary.MemoryProtection.ReadWrite);
                }
                catch (Exception)
                {
                    if (tviPtr == IntPtr.Zero)
                        throw new Exception("Could not allocate memory in owning process of TreeView", new Win32Exception());
                }

                try
                {
                    IntPtr textPtr = IntPtr.Add(tviPtr, (int)tviSize);

                    PInvokeLibrary.TVITEM tvi = new PInvokeLibrary.TVITEM();
                    tvi.mask = (uint)WindowsMessages.TVIF_TEXT;
                    tvi.hItem = hItem;
                    tvi.cchTextMax = (int)WindowsMessages.MY_MAXLVITEMTEXT;
                    tvi.pszText = textPtr;

                    IntPtr ptr = Marshal.AllocHGlobal((IntPtr)tviSize);
                    try
                    {
                        IntPtr myPtr = IntPtr.Zero;
                        Marshal.StructureToPtr(tvi, ptr, false);
                        if (!PInvokeLibrary.WriteProcessMemory(process, tviPtr, ptr, (int)tviSize, out myPtr))
                            throw new Exception("Could not write to memory in owning process of TreeView", new Win32Exception());
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }

                    if ((int)PInvokeLibrary.SendMessage(treeViewHwnd, (int)WindowsMessages.TVM_GETITEM, 0, tviPtr) != 1)
                        throw new Exception("Could not get item data from TreeView");

                    ptr = Marshal.AllocHGlobal((int)textSize);
                    try
                    {
                        int bytesRead = 0;
                        IntPtr bytesReadPtr = IntPtr.Zero;
                        if (!PInvokeLibrary.ReadProcessMemory(process, textPtr, ptr, (int)textSize, out bytesReadPtr))
                            throw new Exception("Could not read from memory in owning process of TreeView", new Win32Exception());

                        bytesRead = (int)bytesReadPtr;

                        if (isUnicode)
                            itemText = Marshal.PtrToStringUni(ptr, bytesRead / 2);
                        else
                            itemText = Marshal.PtrToStringAnsi(ptr, bytesRead);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }
                }
                finally
                {
                    PInvokeLibrary.VirtualFreeEx(process, tviPtr, 0, (PInvokeLibrary.AllocationType)PInvokeLibrary.FreeType.Release);
                }
            }
            finally
            {
                PInvokeLibrary.CloseHandle(process);
            }

            //char[] arr = itemText.ToCharArray(); //<== use this array to look at the bytes in debug mode

            return itemText;
        }

        public List<string> ExtractWindowTextByHandle(IntPtr handle)
        {
            var extractedText = new List<string>();
            List<IntPtr> childObjects = new WindowHandleInfo(handle).GetAllChildHandles();
            for (int i = 0; i < childObjects.Count; i++)
            {
                extractedText.Add(GetControlText(childObjects[i]));
            }
            return extractedText;
        }

        public List<string> ExtractTextByProcessName(string handle)
        {
            List<System.IntPtr> childObjects = new List<System.IntPtr>();
            var extractedText = new List<string>();
            Process[] anotherApps = Process.GetProcessesByName("s7tgtopx");
            if (anotherApps.Length > 0)
            {
                if (anotherApps[0] != null)
                {
                    childObjects = new WindowHandleInfo(anotherApps[0].MainWindowHandle).GetAllChildHandles();
                    for (int i = 0; i < childObjects.Count; i++)
                    {
                        extractedText.Add(GetControlText(childObjects[i]));
                    }
                }
            }
            return extractedText;
        }
        #endregion

        private void SleepUntilDownloadFeedback(int clientIndex)
        {
            bool finished = false;
            bool copied = false;
            DateTime copiedTime = DateTime.Now;
            DateTime comparTime = copiedTime;

            var filePath = pathTextBox.Text + "(" + clientIndex + ")\\winccom\\LOAD.LOG";

            System.Threading.Thread.Sleep(30000);

            while (finished == false)
            {
                if (File.Exists(filePath))
                {
                    System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                    System.IO.StreamReader sr = new System.IO.StreamReader(fs);
                    List<String> lst = new List<string>();

                    while (!sr.EndOfStream)
                        lst.Add(sr.ReadLine());

                    var lastLine = lst.Last(); //File.ReadLines(filePath).Last();
                    finished = lastLine.Contains("The lock on the project was removed"); //now can click ok to conclude download and onto next

                    //safety wait for 40 seconds after the files have been copied, then check is finished...
                    if (lastLine.Contains("The files were copied successfully") && copiedTime == comparTime)
                    {
                        copied = true;
                        copiedTime = DateTime.Now;
                    }
                    if (lastLine.Contains("The computer name was changed in the project") && copied == true)
                    {
                        int diffInSeconds = (DateTime.Now - copiedTime).Seconds;
                        if (finished == false && diffInSeconds > 40)
                        {
                            finished = true;
                        }
                    }
                }
                System.Threading.Thread.Sleep(500);
            }
        }

        private static List<IntPtr> FindWindowByType(IntPtr hWndParent, string type) //something's off
        {
            var list = new List<IntPtr>();
            int ct = 0;
            IntPtr result = IntPtr.Zero;

            do
            {
                result = PInvokeLibrary.FindWindowEx(hWndParent, IntPtr.Zero, type, null);
                if (result != IntPtr.Zero)
                {
                    list.Add(result);
                    ++ct;
                }
            } while (result != IntPtr.Zero);

            return list;
        }

        private void GetIpsLmHosts()
        {
            ipList = new List<string>();
            sdList = new List<string>();
            string lmhostPath = ipTextBox.Text;
            if (File.Exists(lmhostPath))
            {
                var lmHosts = File.ReadLines(lmhostPath);
                foreach (var item in lmHosts)
                {
                    if (item.IndexOf("HMIC") > 0
                        || item.IndexOf("HmiC") > 0
                        || ((item.IndexOf("HmiE01") < 0 && item.IndexOf("HmiE0") > 0)
                        || (item.IndexOf("HMIE01") < 0 && item.IndexOf("HMIE0") > 0))
                        && item.StartsWith("#") == false && item != "")
                    {
                        ipList.Add(item);
                    }
                    if ((item.IndexOf("HMID01") > 0 ||
                        item.IndexOf("HmiD01") > 0 ||
                        item.IndexOf("HMIE01") > 0 ||
                        item.IndexOf("HmiE01") > 0 ||
                        item.IndexOf("HmiS") > 0 ||
                        item.IndexOf("HMIS") > 0) &&
                        item.StartsWith("#") == false && item != "")
                    {
                        sdList.Add(item);
                    }
                }
            }
        }

        private void KeepConfig()
        {
            string configPath = "NCM_Downloader.ini";
            var configFile = File.CreateText(configPath);
            configFile.WriteLine(pathTextBox.Text);
            configFile.WriteLine(firstClientIndexBox.Text);
            configFile.WriteLine(numClTextBox.Text);
            configFile.WriteLine(ipTextBox.Text);
            configFile.WriteLine(unTextBox.Text);
            configFile.WriteLine(passTextBox.Text);
            configFile.WriteLine(rdpCheckBox.Checked);
            configFile.WriteLine(widthBox.Text);
            configFile.WriteLine(heightBox.Text);
            configFile.WriteLine(leftBox.Text);
            configFile.WriteLine(topBox.Text);
            configFile.WriteLine("");
            configFile.WriteLine("Authored by Muresan Radu-Adrian (MURA02)");
            configFile.Close();
        }

        private void SetTooltips()
        {
            var toolTip1 = new System.Windows.Forms.ToolTip();
            toolTip1.SetToolTip(widthBox, "Set RDP Window Width");

            var toolTip2 = new System.Windows.Forms.ToolTip();
            toolTip2.SetToolTip(heightBox, "Set RDP Window Height");

            var toolTip3 = new System.Windows.Forms.ToolTip();
            toolTip3.SetToolTip(leftBox, "Set RDP Window x");

            var toolTip4 = new System.Windows.Forms.ToolTip();
            toolTip4.SetToolTip(topBox, "Set RDP Window y");

            var toolTip5 = new System.Windows.Forms.ToolTip();
            toolTip5.SetToolTip(ipTextBox, "Set path to your lmhosts file.");

            var toolTip6 = new System.Windows.Forms.ToolTip();
            toolTip6.SetToolTip(ipTextBox, "Set path to your lmhosts file.");

            var toolTip7 = new System.Windows.Forms.ToolTip();
            toolTip7.SetToolTip(pathTextBox, "Set path to your LOAD.LOG files as specified in readme.");
        }

        // Send a series of key presses to the application.
        private const int IDOK = 1;

        private void Button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            statusLabel.Text = "Progress: ";
            //
            // check inputs
            //
            if (pathTextBox.Text == null || pathTextBox.Text == "")
            {
                MessageBox.Show(new Form { TopMost = true }, " Please input the path in the following form: " + @"D:\Project\SDIB_TCM\wincproj\SDIB_TCM_CLT_Ref");
                return;
            }
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                //MessageBox.Show(new Form { TopMost = true }, " You have not checked any clients to download to!");
                statusLabel.Text = "Status: You have not checked any clients to download to!";
                statusLabel.Refresh();
                return;
            }
            KeepConfig();
            DownloadProcess();
        }

        private static void LogToFile(string content)
        {
            using (var fileWriter = new StreamWriter(Application.StartupPath + "\\NCM_Downloader.logger", true))
            {
                DateTime date = DateTime.UtcNow;
                fileWriter.WriteLine(date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content);
                fileWriter.Close();
            }
        }

        private void DownloadProcess()
        {
            int clientProg = progressBar1.Maximum / checkedListBox1.CheckedIndices.Count;
            // init logFile
            //
            LogToFile("Started actions ******************************************");
            //
            // initialize handles for windows
            //
            var ncmWndClass = "s7tgtopx"; //ncm manager main window
            var anyPopupClass = "#32770"; //usually any popup

            IntPtr ncmHandle = PInvokeLibrary.FindWindow(ncmWndClass, null);
            IntPtr tgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, null);

            //Testing(ncmHandle, anyPopupClass);

            //richTextBox1.Text = Convert.ToString(childNode);

            if (ncmHandle == IntPtr.Zero)
            {
                MessageBox.Show(new Form { TopMost = true }, " Simatic NCM Manager is not running.");
                LogToFile("NCM Manager was not running.");

                return;
            }

            //handle the missing software package notification
            if (tgtWndHandle != IntPtr.Zero)
            {
                IntPtr btnHandle = PInvokeLibrary.FindWindowEx(tgtWndHandle, IntPtr.Zero, "Button", null);

                if (btnHandle != IntPtr.Zero)
                {
                    PInvokeLibrary.SetForegroundWindow(tgtWndHandle);
                    PInvokeLibrary.SendMessage(btnHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                }
            }
            else
            {
                LogToFile("The missing software package notification did not appear - that is ok");

                PInvokeLibrary.SetForegroundWindow(ncmHandle);
            }

            if (tgtWndHandle != IntPtr.Zero)
            {
                //now perform actions from now on, i.e. CTRL+L
                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    if (i > 0)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    // sys tree view 32 already selected when focusing - navigate from here
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    System.Threading.Thread.Sleep(200);
                    int clientIndex = checkedListBox1.CheckedIndices[i];
                    string clientName = checkedListBox1.CheckedItems[i].ToString();
                    //
                    //get ip here
                    //
                    var myIp = ipList.Where(x => x.Contains(clientName)).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                    //
                    //download process starts here - first needs to navigate to correct index
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    ResetExpansions(ncmHandle);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    ReturnToFirstClient(ncmHandle);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    DownloadToCurrentIndex(clientIndex, ncmHandle);

                    LogToFile("Attempting download to client " + clientName);


                    var started = DateTime.Now;

                    //now new window with download os
                    System.Threading.Thread.Sleep(500);
                    IntPtr osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Download OS");
                    if (osDldTgtWndHandle == IntPtr.Zero)
                    {
                        osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                        if (osDldTgtWndHandle == IntPtr.Zero)
                        {
                            osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, null);
                        }
                    }

                    if (osDldTgtWndHandle != IntPtr.Zero)
                    {
                        PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                        IntPtr DlButtonHandle = PInvokeLibrary.FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "OK");
                        if (DlButtonHandle != IntPtr.Zero)
                        {
                            PInvokeLibrary.SendMessage(DlButtonHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                            System.Threading.Thread.Sleep(500); //important to wait a bit

                            IntPtr deactivateRTPopup = PInvokeLibrary.FindWindow(anyPopupClass, "Target system");
                            if (deactivateRTPopup != IntPtr.Zero)
                            {
                                PInvokeLibrary.SetForegroundWindow(deactivateRTPopup);
                                System.Threading.Thread.Sleep(500); //important to wait a bit
                                IntPtr YesButton = PInvokeLibrary.FindWindowEx(deactivateRTPopup, IntPtr.Zero, "Button", "&Yes");
                                if (YesButton != IntPtr.Zero)
                                {
                                    PInvokeLibrary.SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                }
                                else
                                {
                                    LogToFile("The Ok Button was not found in the deactivation window!");

                                    //MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                }
                            }
                            else
                            {
                                LogToFile("Did not find target system popup! (runtime was not active on this client)");

                                //MessageBox.Show(new Form { TopMost = true }, "Did not find target system popup!"); //careful to focus on it
                            }

                            #region TaskKill process if deactivating/closing

                            var processName = "CCOnScreenKeyboard";
                            var user = unTextBox.Text;
                            var pass = passTextBox.Text;
                            var ip = myIp;
                            IntPtr killGuideHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                            if (killGuideHandle != IntPtr.Zero)
                            {
                                //Download to target system was completed successfully. do not send enter until this text is present in the window...
                                bool flagKilled = false;
                                do
                                {
                                    var killGuideText = ExtractWindowTextByHandle(killGuideHandle);
                                    if (killGuideText.Where(x => x.Contains("Closing project on the Runtime OS")).Count() > 0 || killGuideText.Where(x => x.Contains("Deactivating project on the Runtime OS")).Count() > 0) //check if closing project takes too long...
                                    {
                                        LogToFile("Attempting to kill " + processName + " at " + ip + " with username " + user + " and password " + pass + " on client " + clientName);
                                        KillProcessViaPowershellOnMachine(clientName, processName);
                                        flagKilled = true;
                                    }
                                    System.Threading.Thread.Sleep(50);
                                } while (flagKilled == false);
                            }

                            #endregion

                            //call remote desktop open here
                            if (rdpCheckBox.Checked == true)
                            {
                                //certificate needs to be generated here
                                System.Threading.Thread.Sleep(10000); //important to wait a bit
                                OpenRemoteSession(myIp, unTextBox.Text, passTextBox.Text);
                            }

                            //if Canceled by the user in LOAD.LOG , assume that RT station not obtainable //read load.log here to find canceled by user
                            PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                            //var downloadTargetWindowText = ExtractWindowTextByHandle(osDldTgtWndHandle);

                            var filePath = pathTextBox.Text + "(" + checkedListBox1.CheckedIndices[i] + 1 + ")\\winccom\\LOAD.LOG";
                            if (File.Exists(filePath))
                            {
                                System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                                System.IO.StreamReader sr = new System.IO.StreamReader(fs);
                                List<String> lst = new List<string>();

                                while (!sr.EndOfStream)
                                    lst.Add(sr.ReadLine());
                            }

                            //SleepUntilDownloadFeedback(clientIndex + 1);
                            System.Threading.Thread.Sleep(30000);

                            IntPtr dldingTgtHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                            if (dldingTgtHandle != IntPtr.Zero)
                            {
                                //Download to target system was completed successfully. do not send enter until this text is present in the window...
                                var dldingTgtText = ExtractWindowTextByHandle(dldingTgtHandle);

                                DateTime foundTime = DateTime.Now;
                                bool stillClosing = false;
                                do
                                {
                                    dldingTgtText = ExtractWindowTextByHandle(dldingTgtHandle);
                                    if (dldingTgtText.Where(x => x.Contains("Closing project on the Runtime OS.")).Count() > 0) //check if closing project takes too long...
                                    {
                                        int diffInSeconds = (DateTime.Now - foundTime).Seconds;
                                        if (stillClosing == false && diffInSeconds > 60)
                                        {
                                            stillClosing = true; //not used yet

                                            MessageBox.Show(new Form { TopMost = true }, "Please check the client " + clientName + ", it seems something is open and prevents project close.");
                                        }
                                    }

                                    if (dldingTgtText.Where(x => x.Contains("Error")).Count() > 0)
                                    {
                                        System.Threading.Thread.Sleep(500);
                                        PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                                        System.Threading.Thread.Sleep(500);
                                        IntPtr YesButton = PInvokeLibrary.FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "Ok");
                                        if (YesButton != IntPtr.Zero)
                                        {
                                            PInvokeLibrary.SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                        }
                                        else
                                        {
                                            LogToFile("The Ok Button was not found in the downloading window!");

                                            MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                        }
                                        LogToFile("Error on download to client " + clientIndex + 1);

                                        continue;
                                    }

                                    if (dldingTgtText.Where(x => x.Contains("Canceled:")).Count() > 0)
                                    {
                                        LogToFile("Client " + clientIndex + " download canceled - RT station not obtainable");

                                        MessageBox.Show(new Form { TopMost = true }, " Client " + clientIndex + " download canceled - RT station not obtainable - will continue to next client download");
                                        continue;
                                    }

                                    System.Threading.Thread.Sleep(1000);

                                    foreach (var c in dldingTgtText)
                                    {
                                        LogToFile(c);
                                    }

                                } while (dldingTgtText.Where(x => x.Contains("Download to target system was completed successfully")).Count() == 0);

                                ClickButtonUsingMessage(dldingTgtHandle, "OK", "Download to target system was completed successfully");
                            }
                        }

                        if (rdpCheckBox.Checked == true)
                        {
                            CloseRemoteSession(myIp);
                        }

                        progressBar1.Value = clientProg * (i + 1);
                        statusLabel.Text = "Progress: " + progressBar1.Value.ToString() + "%";
                        var ended = DateTime.Now;
                        var secElapsed = Math.Round((ended - started).TotalSeconds, 2);
                        LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() +
                            " : Finished download process for machine " + checkedListBox1.CheckedItems[i].ToString() +
                            " in " + secElapsed.ToString() + " seconds");

                        //int currentElapsed = Convert.ToInt32(label9.Text.Split(Convert.ToChar("~"))[1]);
                        //label9.Text = "Remaining [s]: ~" + (currentElapsed - 100).ToString();
                        label9.Refresh();
                        progressBar1.Refresh();
                        statusLabel.Refresh();
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Could not focus on download popup!"); //careful to focus on it
                    }
                }
                LogToFile("Closing logfile");

                MessageBox.Show(new Form { TopMost = true }, "The NCM download process has been finished!"); //careful to focus on it
            }
        }

        private static void TaskKill(string user, string pass, string ip, string processName)
        {
            var connectoptions = new ConnectionOptions();
            connectoptions.Username = user; //@"YourDomainName\UserName";
            connectoptions.Password = pass; //"User Password";

            string ipAddress = ip; //"192.168.206.53";
            ManagementScope scope = new ManagementScope(@"\\" + ipAddress + @"\root\cimv2", connectoptions);

            // WMI query
            var query = new SelectQuery("select * from Win32_process where name = '" + processName + "'");

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                foreach (ManagementObject process in searcher.Get()) // this is the fixed line
                {
                    process.InvokeMethod("Terminate", null);
                }
            }
            Console.ReadLine();
        }

        private void ClickButtonUsingMessage(IntPtr windowHandle, string buttonText, string windowText)
        {
            bool success = true;
            do
            {
                try
                {
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
                    IntPtr OkButton = PInvokeLibrary.FindWindowEx(windowHandle, IntPtr.Zero, "Button", buttonText);
                    if (OkButton != IntPtr.Zero)
                    {
                        PInvokeLibrary.SendMessage(OkButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, OkButton);
                        var checkText = ExtractWindowTextByHandle(windowHandle);
                        if (checkText.Where(x => x.Contains(windowText)).Count() == 0)
                        {
                            success = true;
                        }
                    }
                    else
                    {
                        LogToFile("The " + buttonText + " button was not found in the downloading window to confirm finish!");

                        //MessageBox.Show(new Form { TopMost = true }, "The OK Button was not found in the downloading to target system window!"); //careful to focus on it
                        success = false;
                    }
                }
                catch (Exception exc)
                {
                    LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + exc.Message);

                    //MessageBox.Show(new Form { TopMost = true }, exc.Message);
                    success = false;
                }
            } while (success == false);
        }

        private void CloseConflictingProcesses(string substr)
        {
            List<Process> myProcesses = new List<Process>();
            //int i = 0;
            foreach (var proc in Process.GetProcesses())
            {
                if (proc.ProcessName.StartsWith(substr))
                {
                    myProcesses.Add(proc);
                }
            }
            foreach (var proc in myProcesses)
            {
                if (proc.Id != Process.GetCurrentProcess().Id)
                {
                    proc.Kill();
                }
            }
        }


        private void OpenRemoteSession(string ip, string un, string pw)
        {
            Process rdcProcess = new Process();
            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
            rdcProcess.StartInfo.Arguments = "/generic:TERMSRV/" + ip + " /user:" + un + " /pass:" + pw;
            rdcProcess.Start();
            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
            rdcProcess.StartInfo.Arguments = "/v " + ip; // ip or name of computer to connect
            rdcProcess.Start();

            IntPtr myRdp;
            do
            {
                myRdp = PInvokeLibrary.FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
                if (myRdp != IntPtr.Zero)
                {
                    // Set the window's position.
                    int width = int.Parse(widthBox.Text);
                    int height = int.Parse(heightBox.Text);
                    int x = int.Parse(leftBox.Text);
                    int y = int.Parse(topBox.Text);

                    // Prepare the WINDOWPLACEMENT structure.
                    var placement = new PInvokeLibrary.WINDOWPLACEMENT();
                    placement.Length = Marshal.SizeOf(placement);

                    // Get the window's current placement.
                    PInvokeLibrary.GetWindowPlacement(myRdp, ref placement);
                    if (placement.ShowCmd != PInvokeLibrary.ShowWindowCommands.Normal)
                    {
                        //alter the placement
                        placement.ShowCmd = PInvokeLibrary.ShowWindowCommands.Normal;
                        //set the changes
                        PInvokeLibrary.SetWindowPlacement(myRdp, ref placement);
                    }

                    PInvokeLibrary.SetWindowPos(myRdp, IntPtr.Zero, x, y, width, height, 0);

                }
                else
                {
                    System.Threading.Thread.Sleep(1000);
                }
            } while (myRdp == IntPtr.Zero);
        }

        private void CloseRemoteSession(string ip)
        {
            IntPtr myRdp = PInvokeLibrary.FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
            if (myRdp != IntPtr.Zero)
            {
                PInvokeLibrary.SendMessage(myRdp, (int)WindowsMessages.WM_CLOSE, (int)IntPtr.Zero, IntPtr.Zero);
                System.Threading.Thread.Sleep(1000);
            }
        }

        private void SendKeyHandled(IntPtr windowHandle, string key)
        {
            bool success;
            do
            {
                try
                {
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
                    SendKeys.Send(key);
                    success = true;
                }
                catch (Exception e)
                {
                    LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + e.Message);

                    success = false;
                    System.Threading.Thread.Sleep(10);
                }
            } while (success == false);
        }

        private void NavigateToIndex(int index, IntPtr ncm)
        {
            for (int i = 0; i < index; i++)
            {
                SendKeyHandled(ncm, "{DOWN}");
            }
        }

        private void ResetExpansions(IntPtr ncm)
        {
            int num = int.Parse(numClTextBox.Text) * 3;
            for (int i = 0; i < num; i++) //go up
            {
                SendKeyHandled(ncm, "{UP}");
            }
            for (int i = 0; i < num; i++) //expand all
            {
                SendKeyHandled(ncm, "{DOWN}");
                for (int j = 0; j < 6; j++)
                {
                    SendKeyHandled(ncm, "{RIGHT}");
                }
            }
            for (int i = 0; i < num; i++) //back to compact
            {
                for (int j = 0; j < 4; j++)
                {
                    SendKeyHandled(ncm, "{LEFT}");
                }
                SendKeyHandled(ncm, "{UP}");
            }
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{DOWN}"); //go to first dev station or whatever in the list
        }

        private void DownloadToCurrentIndex(int index, IntPtr ncm)
        {
            NavigateToIndex(index, ncm);
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "^(l)");
            System.Threading.Thread.Sleep(500);
        }

        private void ReturnToFirstClient(IntPtr ncm)
        {
            for (int i = 0; i < 10; i++)
            {
                SendKeyHandled(ncm, "{LEFT}");
            }
            SendKeyHandled(ncm, "{RIGHT}");

            for (int i = 0; i < int.Parse(firstClientIndexBox.Text); i++)
            {
                SendKeyHandled(ncm, "{DOWN}");
            }
        }

        // Send a key to the button when the user double-clicks anywhere on the form.
        private void Form1_DoubleClick(object sender, EventArgs e)
        {
            // Send the enter key to the button, which raises the click
            // event for the button. This works because the tab stop of
            // the button is 0.
            SendKeys.Send("{ENTER}");
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

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NCMForm));
            this.label1 = new System.Windows.Forms.Label();
            this.firstClientIndexBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.numClTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.pathTextBox = new System.Windows.Forms.TextBox();
            this.ipTextBox = new System.Windows.Forms.TextBox();
            this.unTextBox = new System.Windows.Forms.TextBox();
            this.passTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.rdpCheckBox = new System.Windows.Forms.CheckBox();
            this.rdpBox1 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.topBox = new System.Windows.Forms.TextBox();
            this.leftBox = new System.Windows.Forms.TextBox();
            this.heightBox = new System.Windows.Forms.TextBox();
            this.widthBox = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.statusLabel = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdpBox1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-71, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Clients start index:";
            this.label1.Visible = false;
            // 
            // firstClientIndexBox
            // 
            this.firstClientIndexBox.Enabled = false;
            this.firstClientIndexBox.Location = new System.Drawing.Point(29, 84);
            this.firstClientIndexBox.Name = "firstClientIndexBox";
            this.firstClientIndexBox.Size = new System.Drawing.Size(28, 20);
            this.firstClientIndexBox.TabIndex = 1;
            this.firstClientIndexBox.Text = "3";
            this.firstClientIndexBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.firstClientIndexBox.Visible = false;
            // 
            // button1
            // 
            this.button1.AutoSize = true;
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Location = new System.Drawing.Point(12, 348);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(300, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Download to selected NCM Clients";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.SystemColors.Control;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(68, 65);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.ScrollAlwaysVisible = true;
            this.checkedListBox1.Size = new System.Drawing.Size(244, 139);
            this.checkedListBox1.TabIndex = 8;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // numClTextBox
            // 
            this.numClTextBox.Enabled = false;
            this.numClTextBox.Location = new System.Drawing.Point(29, 113);
            this.numClTextBox.Name = "numClTextBox";
            this.numClTextBox.Size = new System.Drawing.Size(28, 20);
            this.numClTextBox.TabIndex = 9;
            this.numClTextBox.Text = "4";
            this.numClTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numClTextBox.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(-71, 116);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Number of clients:";
            this.label5.Visible = false;
            // 
            // pathTextBox
            // 
            this.pathTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.pathTextBox.Location = new System.Drawing.Point(68, 11);
            this.pathTextBox.Name = "pathTextBox";
            this.pathTextBox.Size = new System.Drawing.Size(244, 18);
            this.pathTextBox.TabIndex = 11;
            this.pathTextBox.Text = "D:\\Project\\SDIB_TCM\\wincproj\\SDIB_TCM_CLT_Ref";
            // 
            // ipTextBox
            // 
            this.ipTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.ipTextBox.Location = new System.Drawing.Point(68, 35);
            this.ipTextBox.Name = "ipTextBox";
            this.ipTextBox.Size = new System.Drawing.Size(244, 18);
            this.ipTextBox.TabIndex = 12;
            this.ipTextBox.Text = "C:\\Windows\\System32\\drivers\\etc\\lmhosts";
            this.ipTextBox.TextChanged += new System.EventHandler(this.ipTextBox_TextChanged);
            // 
            // unTextBox
            // 
            this.unTextBox.Location = new System.Drawing.Point(56, 19);
            this.unTextBox.Name = "unTextBox";
            this.unTextBox.Size = new System.Drawing.Size(238, 20);
            this.unTextBox.TabIndex = 13;
            this.unTextBox.Text = "SDI";
            // 
            // passTextBox
            // 
            this.passTextBox.Location = new System.Drawing.Point(56, 45);
            this.passTextBox.Name = "passTextBox";
            this.passTextBox.Size = new System.Drawing.Size(238, 20);
            this.passTextBox.TabIndex = 14;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label2.Location = new System.Drawing.Point(10, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "LmHosts";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label3.Location = new System.Drawing.Point(6, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "User";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label4.Location = new System.Drawing.Point(6, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 17;
            this.label4.Text = "Pass";
            // 
            // rdpCheckBox
            // 
            this.rdpCheckBox.AutoSize = true;
            this.rdpCheckBox.Location = new System.Drawing.Point(13, 217);
            this.rdpCheckBox.Name = "rdpCheckBox";
            this.rdpCheckBox.Size = new System.Drawing.Size(79, 17);
            this.rdpCheckBox.TabIndex = 18;
            this.rdpCheckBox.Text = "Start RDPs";
            this.rdpCheckBox.UseVisualStyleBackColor = true;
            // 
            // rdpBox1
            // 
            this.rdpBox1.Controls.Add(this.label8);
            this.rdpBox1.Controls.Add(this.label7);
            this.rdpBox1.Controls.Add(this.topBox);
            this.rdpBox1.Controls.Add(this.leftBox);
            this.rdpBox1.Controls.Add(this.heightBox);
            this.rdpBox1.Controls.Add(this.widthBox);
            this.rdpBox1.Controls.Add(this.label4);
            this.rdpBox1.Controls.Add(this.label3);
            this.rdpBox1.Controls.Add(this.unTextBox);
            this.rdpBox1.Controls.Add(this.passTextBox);
            this.rdpBox1.Location = new System.Drawing.Point(12, 239);
            this.rdpBox1.Name = "rdpBox1";
            this.rdpBox1.Size = new System.Drawing.Size(300, 103);
            this.rdpBox1.TabIndex = 19;
            this.rdpBox1.TabStop = false;
            this.rdpBox1.Text = "Remote Desktop Automation";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label8.Location = new System.Drawing.Point(180, 75);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(28, 13);
            this.label8.TabIndex = 23;
            this.label8.Text = "Pos.";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label7.Location = new System.Drawing.Point(6, 75);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(28, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Res.";
            // 
            // topBox
            // 
            this.topBox.Location = new System.Drawing.Point(264, 71);
            this.topBox.Name = "topBox";
            this.topBox.Size = new System.Drawing.Size(30, 20);
            this.topBox.TabIndex = 21;
            this.topBox.Text = "300";
            // 
            // leftBox
            // 
            this.leftBox.Location = new System.Drawing.Point(228, 71);
            this.leftBox.Name = "leftBox";
            this.leftBox.Size = new System.Drawing.Size(30, 20);
            this.leftBox.TabIndex = 20;
            this.leftBox.Text = "700";
            // 
            // heightBox
            // 
            this.heightBox.Location = new System.Drawing.Point(92, 71);
            this.heightBox.Name = "heightBox";
            this.heightBox.Size = new System.Drawing.Size(30, 20);
            this.heightBox.TabIndex = 19;
            this.heightBox.Text = "480";
            // 
            // widthBox
            // 
            this.widthBox.Location = new System.Drawing.Point(56, 71);
            this.widthBox.Name = "widthBox";
            this.widthBox.Size = new System.Drawing.Size(30, 20);
            this.widthBox.TabIndex = 18;
            this.widthBox.Text = "640";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(189, 217);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkBox1.Size = new System.Drawing.Size(117, 17);
            this.checkBox1.TabIndex = 20;
            this.checkBox1.Text = "Select/Deselect All";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label6.Location = new System.Drawing.Point(10, 14);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "NCMPath";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(13, 377);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(299, 13);
            this.progressBar1.TabIndex = 22;
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.BackColor = System.Drawing.Color.Transparent;
            this.statusLabel.Location = new System.Drawing.Point(9, 399);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(68, 13);
            this.statusLabel.TabIndex = 23;
            this.statusLabel.Text = "Progress: 0%";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(213, 377);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(0, 13);
            this.label9.TabIndex = 24;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(241, 17);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(53, 23);
            this.button2.TabIndex = 25;
            this.button2.Text = "Kill";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(102, 20);
            this.textBox1.TabIndex = 26;
            this.textBox1.Text = "TcmHmiC05";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(114, 19);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(121, 20);
            this.textBox2.TabIndex = 27;
            this.textBox2.Text = "Taskmgr";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Location = new System.Drawing.Point(12, 415);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(300, 51);
            this.groupBox1.TabIndex = 28;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Process Control";
            // 
            // NCMForm
            // 
            this.ClientSize = new System.Drawing.Size(324, 474);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.rdpCheckBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.rdpBox1);
            this.Controls.Add(this.ipTextBox);
            this.Controls.Add(this.pathTextBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.numClTextBox);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.firstClientIndexBox);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NCMForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "NCM Manager Automation";
            this.rdpBox1.ResumeLayout(false);
            this.rdpBox1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Checked);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
            checkedListBox1_SelectedIndexChanged(checkBox1, new System.EventArgs());
        }

        private void ipTextBox_TextChanged(object sender, EventArgs e)
        {
            RenewIpsOrInit();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //label9.Text = "Remaining [s]: ~" + (checkedListBox1.CheckedIndices.Count * 100).ToString();
            label9.Refresh();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var id = textBox1.Text;
            KillProcessViaPowershellOnMachine(id, textBox2.Text);

        }

        private void KillTaskInRDP(string ip, string processName)
        {
            IntPtr myRdp = PInvokeLibrary.FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");

            Interoperability.PInvokeLibrary.SetForegroundWindow(myRdp);

            _ = PInvokeLibrary.GetWindowRect(myRdp, out PInvoke.RECT myRdpRect);


            ClickInWindowAtXY(myRdp, ((myRdpRect.left + myRdpRect.right) / 2) /*(Convert.ToInt32(widthBox.Text) / 480 * 9)*/, ((myRdpRect.top + myRdpRect.bottom) / 2)/*+ myRdpRect.bottom - (Convert.ToInt32(heightBox.Text) / 480 * 9)*/, 1);
            LogToFile("Clicked at " + (myRdpRect.left + (Convert.ToInt32(widthBox.Text) / 480 * 9)).ToString() + " width, " + (myRdpRect.top + myRdpRect.bottom - (Convert.ToInt32(heightBox.Text) / 480 * 9)).ToString() + " height");
            System.Threading.Thread.Sleep(500);

            var cmdString = @"Taskkill / IM " + processName + " / F";

            System.Threading.Thread.Sleep(200);
            foreach (var c in cmdString)
            {
                SendKeyHandled(myRdp, c.ToString());
            }

            SendKeyHandled(myRdp, "^({ESC})");
            System.Threading.Thread.Sleep(500);

            SendKeyHandled(myRdp, "c");
            SendKeyHandled(myRdp, "m");
            SendKeyHandled(myRdp, "d");
            System.Threading.Thread.Sleep(200);
            SendKeyHandled(myRdp, "{ENTER}");

            /*var*/
            cmdString = @"Taskkill / IM " + processName + " / F";

            System.Threading.Thread.Sleep(200);
            foreach (var c in cmdString)
            {
                SendKeyHandled(myRdp, c.ToString());
            }

            LogToFile("should have succeeded by now");
        }

        private void KillProcessViaPowershellOnMachine(string machine, string process)
        {
            try
            {
                AddToTrustedHosts(machine);
            }
            catch (Exception e)
            {
                LogToFile(e.Message);
            }
            finally
            {
                //Invoke-command -computername "TcmHmiC05" {Get-Process | ? {$_.name -match 'CCOnScreenKeyboard'} | Stop-Process -Force}
                try
                {
                    Runspace runSpace = RunspaceFactory.CreateRunspace();
                    runSpace.Open();
                    Pipeline pipeline = runSpace.CreatePipeline();

                    Command invokeScript = new Command("Invoke-Command");
                    RunspaceInvoke invoke = new RunspaceInvoke();
                    //Invoke-Command -scriptBlock
                    ScriptBlock sb = invoke.Invoke("{Get-Process | ? {$_.name -match '" + process + "'} | Stop-Process -Force}")[0].BaseObject as ScriptBlock;
                    invokeScript.Parameters.Add("computername", machine);
                    invokeScript.Parameters.Add("scriptBlock", sb);

                    pipeline.Commands.Add(invokeScript);
                    Collection<PSObject> output = pipeline.Invoke();
                    foreach (PSObject obj in output)
                    {
                        LogToFile(obj.ToString());
                    }
                }
                catch (Exception e)
                {
                    LogToFile(e.Message + " at kill process");
                } finally
                {
                    LogToFile("could not kill process");
                }
            }
        }
        private void AddToTrustedHosts(string machine)
        {
            //RunPowershellCommand(machine);
            ////Invoke-command -computername "TcmHmiC05" {Get-Process | ? {$_.name -match 'CCOnScreenKeyboard'} | Stop-Process -Force}
            //try
            //{
            //    Runspace runSpace = RunspaceFactory.CreateRunspace();
            //    runSpace.Open();
            //    Pipeline pipeline = runSpace.CreatePipeline();

            //    Command invokeScript = new Command("Set-Item");
            //    RunspaceInvoke invoke = new RunspaceInvoke();
            //    //Invoke-Command -scriptBlock
            //    //ScriptBlock sb = invoke.Invoke("{Get-Process | ? {$_.name -match '" + process + "'} | Stop-Process -Force}")[0].BaseObject as ScriptBlock;
            //    invokeScript.Parameters.Add("Path[0]", @"WSMan:\localhost\Client\TrustedHosts");
            //    invokeScript.Parameters.Add("Value", machine);

            //    pipeline.Commands.Add(invokeScript);
            //    Collection<PSObject> output = pipeline.Invoke();
            //    foreach (PSObject obj in output)
            //    {
            //        LogToFile(obj.ToString());
            //    }
            //}
            //catch (Exception e)
            //{
            //    LogToFile(e.Message + " at add to trusted hosts");
            //}

            //System.Security.SecureString pass = new System.Security.SecureString();
            //string strPass = "!sysadmin";
            //foreach (char c in strPass)
            //{
            //    pass.AppendChar(c);
            //}

            System.Diagnostics.Process.Start(Application.StartupPath + "\\StartTrustedHosts.ps1");

        }

        private void RunPowershellCommand(string machine)
        {
            try
            {
                using (PowerShell PowerShellInst = PowerShell.Create())
                {
                    //if trying to read from file
                    //string path = Application.StartupPath + "\\AddTrustedHosts.ps1";
                    //if (!string.IsNullOrEmpty(path))
                    PowerShellInst.AddScript(@"Set-Item WSMan:\localhost\Client\TrustedHosts -Value '" + machine + "' -Force");
                    //PowerShellInst.addsta
                    //PowerShellInst.AddCommand("Set-Item");
                    //PowerShellInst.AddParameter(@"WSMan:\localhost\Client\TrustedHosts");
                    //PowerShellInst.AddParameter()


                    Collection<PSObject> PSOutput = PowerShellInst.Invoke();
                    foreach (PSObject obj in PSOutput)
                    {
                        if (obj != null)
                        {
                            LogToFile(obj.Properties["EntryType"].Value.ToString() + " - ");
                            LogToFile(obj.Properties["Source"].Value.ToString() + " - ");
                            LogToFile(obj.Properties["Message"].Value.ToString() + " - ");
                        }
                    }
                    LogToFile("Ran Powershell Set-Item command");
                    //Console.Read();
                }
            }
            catch (Exception e)
            {
                LogToFile(e.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}