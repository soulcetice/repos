using WindowHandle;
using WindowsUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;

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
        private Form frm1 = new Form();
                
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new NCMForm());
            //Done by Muresan Radu-Adrian MURA02 20200716
        }

        #region ImportDlls
        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Define the SetWindowPos API function.
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

        //get objects in window ?
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr handleParent, IntPtr handleChild, string className, string WindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(int hWnd, int Msg, int wparam, int lparam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(ProcessAccessFlags processAccess, bool bInheritHandle, int processId);
        public static IntPtr OpenProcess(Process proc, ProcessAccessFlags flags) { return OpenProcess(flags, false, proc.Id); }

        [DllImport("kernel32.dll", SetLastError = true)]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseHandle(IntPtr hObject);

        [DllImport("user32")]
        public static extern bool IsWindowUnicode(IntPtr hwnd);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, int dwSize, AllocationType dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [Out, MarshalAs(UnmanagedType.AsAny)] object lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [MarshalAs(UnmanagedType.AsAny)] object lpBuffer, int dwSize, out IntPtr lpNumberOfBytesWritten);

        #endregion
        
        #region Flags
        [Flags]
        private enum SetWindowPosFlags : uint
        {
            SynchronousWindowPosition = 0x4000,
            DeferErase = 0x2000,
            DrawFrame = 0x0020,
            FrameChanged = 0x0020,
            HideWindow = 0x0080,
            DoNotActivate = 0x0010,
            DoNotCopyBits = 0x0100,
            IgnoreMove = 0x0002,
            DoNotChangeOwnerZOrder = 0x0200,
            DoNotRedraw = 0x0008,
            DoNotReposition = 0x0200,
            DoNotSendChangingEvent = 0x0400,
            IgnoreResize = 0x0001,
            IgnoreZOrder = 0x0004,
            ShowWindow = 0x0040,
        }
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

        public enum AllocationType
        {
            Commit = 0x1000,
            Reserve = 0x2000,
            Decommit = 0x4000,
            Release = 0x8000,
            Reset = 0x80000,
            Physical = 0x400000,
            TopDown = 0x100000,
            WriteWatch = 0x200000,
            LargePages = 0x20000000
        }

        public enum MemoryProtection
        {
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            NoAccess = 0x01,
            ReadOnly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            GuardModifierflag = 0x100,
            NoCacheModifierflag = 0x200,
            WriteCombineModifierflag = 0x400
        }

        struct TVITEM
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
        }

        public enum TreeViewMsg
        {
            BN_CLICKED = 0xf5,
            TV_CHECKED = 0x2000,
            TV_FIRST = 0x1100,
            TVGN_ROOT = 0x0,
            TVGN_NEXT = 0x1,
            TVGN_CHILD = 0x4,
            TVGN_FIRSTVISIBLE = 0x5,
            TVGN_NEXTVISIBLE = 0x6,
            TVGN_CARET = 0x9,
            TVM_SELECTITEM = (TV_FIRST + 11),
            TVM_GETNEXTITEM = (TV_FIRST + 10),
            TVM_GETITEM = (TV_FIRST + 12),
            TVIF_TEXT = 0x1
        }
        public enum FreeType
        {
            Decommit = 0x4000,
            Release = 0x8000,
        }
        #endregion

        private List<string> ipList = new List<string>();
        private List<string> sdList = new List<string>();

        public NCMForm()
        {
            CloseConflictingProcesses("NCM_Downloader");

            InitializeComponent();

            //textBox4_TextChanged(numClTextBox, new System.EventArgs());

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
            Int32 titleSize = SendMessage((int)hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, 0).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

            return title.ToString();
        }

        #region TextExtractArea
        public static string GetTreeItemText(IntPtr treeViewHwnd, IntPtr hItem)
        {
            string itemText;

            uint pid;
            GetWindowThreadProcessId(treeViewHwnd, out pid);

            IntPtr process = OpenProcess(ProcessAccessFlags.VirtualMemoryOperation | ProcessAccessFlags.VirtualMemoryRead | ProcessAccessFlags.VirtualMemoryWrite | ProcessAccessFlags.QueryInformation, false, (int)pid);
            if (process == IntPtr.Zero)
                throw new Exception("Could not open handle to owning process of TreeView", new Win32Exception());

            try
            {
                uint tviSize = (uint)Marshal.SizeOf(typeof(TabItem));

                uint textSize = (uint)WindowsUtilities.WindowsMessages.MY_MAXLVITEMTEXT;
                bool isUnicode = IsWindowUnicode(treeViewHwnd);
                if (isUnicode)
                    textSize *= 2;

                IntPtr tviPtr = VirtualAllocEx(process, IntPtr.Zero, tviSize + textSize, AllocationType.Commit, MemoryProtection.ReadWrite);
                if (tviPtr == IntPtr.Zero)
                    throw new Exception("Could not allocate memory in owning process of TreeView", new Win32Exception());

                try
                {
                    IntPtr textPtr = IntPtr.Add(tviPtr, (int)tviSize);

                    TVITEM tvi = new TVITEM();
                    tvi.mask = (uint)WindowsMessages.TVIF_TEXT;
                    tvi.hItem = hItem;
                    tvi.cchTextMax = (int)WindowsMessages.MY_MAXLVITEMTEXT;
                    tvi.pszText = textPtr;

                    IntPtr ptr = Marshal.AllocHGlobal((IntPtr)tviSize);
                    try
                    {
                        Marshal.StructureToPtr(tvi, ptr, false);
                        if (!WriteProcessMemory(process, tviPtr, ptr, (int)tviSize, out IntPtr bytesWritten))
                            throw new Exception("Could not write to memory in owning process of TreeView", new Win32Exception());
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }

                    if (SendMessage(treeViewHwnd, (int)(isUnicode ? WindowsMessages.TVM_GETITEMW : WindowsMessages.TVM_GETITEMA), 0, tviPtr) != IntPtr.Zero)
                        throw new Exception("Could not get item data from TreeView");

                    ptr = Marshal.AllocHGlobal((IntPtr)textSize);
                    try
                    {
                        IntPtr bytesRead;
                        if (!ReadProcessMemory(process, textPtr, ptr, (int)textSize, out bytesRead))
                            throw new Exception("Could not read from memory in owning process of TreeView", new Win32Exception());

                        if (isUnicode)
                            itemText = Marshal.PtrToStringUni(ptr, (int)bytesRead / 2);
                        else
                            itemText = Marshal.PtrToStringAnsi(ptr, (int)bytesRead);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }
                }
                finally
                {
                    VirtualFreeEx(process, tviPtr, 0, (AllocationType)FreeType.Release);
                }
            }
            finally
            {
                CloseHandle(process);
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
                result = FindWindowEx(hWndParent, IntPtr.Zero, type, null);
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
            //
            //write config file
            //
            KeepConfig();
            //
            //
            DownloadProcess();
        }

        private void DownloadProcess()
        {
            int clientProg = progressBar1.Maximum / checkedListBox1.CheckedIndices.Count;
            // init logFile
            //
            string logPath = Application.StartupPath + "\\NCM_Downloader.logger";
            var logFile = File.CreateText(logPath);
            logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Started actions");
            logFile.Flush();
            //
            // initialize handles for windows
            //
            var ncmWndClass = "s7tgtopx"; //ncm manager main window
            var anyPopupClass = "#32770"; //usually any popup

            IntPtr ncmHandle = FindWindow(ncmWndClass, null);
            IntPtr tgtWndHandle = FindWindow(anyPopupClass, null);

            //IntPtr trHandle = FindWindowEx(ncmHandle, IntPtr.Zero, "SysTreeView32", null);
            //IntPtr treeH = new IntPtr(0x000204F2);

            //var childNode = SendMessage(treeH, (int)WindowsMessages.TVM_GETITEMHEIGHT, 0, IntPtr.Zero); //works
            //var childNode = SendMessage(treeH, (int)WindowsMessages.TVM_GETCOUNT, 0, IntPtr.Zero);

            //richTextBox1.Text = Convert.ToString(childNode);

            if (ncmHandle == IntPtr.Zero)
            {
                MessageBox.Show(new Form { TopMost = true }, " Simatic NCM Manager is not running.");
                logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : NCM Manager was not running.");
                logFile.Flush();
                return;
            }

            //handle the missing software package notification
            if (tgtWndHandle != IntPtr.Zero)
            {
                IntPtr btnHandle = FindWindowEx(tgtWndHandle, IntPtr.Zero, "Button", null);

                if (btnHandle != new IntPtr(0x00000000))
                {
                    SetForegroundWindow(tgtWndHandle);
                    SendMessage(btnHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                    SetForegroundWindow(ncmHandle);
                }
            }
            else
            {
                logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : The missing software package notification did not appear - that is ok");
                logFile.Flush();
                SetForegroundWindow(ncmHandle);
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
                    SetForegroundWindow(ncmHandle);
                    System.Threading.Thread.Sleep(500);
                    int clientIndex = checkedListBox1.CheckedIndices[i];
                    string clientName = checkedListBox1.CheckedItems[i].ToString();
                    //
                    //get ip here
                    //
                    var myIp = ipList.Where(x => x.Contains(clientName)).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                    //
                    //download process starts here - first needs to navigate to correct index
                    SetForegroundWindow(ncmHandle);
                    ResetExpansions(ncmHandle, logFile);
                    SetForegroundWindow(ncmHandle);
                    ReturnToFirstClient(ncmHandle, logFile);
                    SetForegroundWindow(ncmHandle);
                    DownloadToCurrentIndex(clientIndex, ncmHandle, logFile);

                    logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Attempting download to client " + clientName);
                    logFile.Flush();

                    var started = DateTime.Now;

                    //now new window with download os
                    System.Threading.Thread.Sleep(500);
                    IntPtr osDldTgtWndHandle = FindWindow(anyPopupClass, "Download OS");
                    if (osDldTgtWndHandle == IntPtr.Zero)
                    {
                        osDldTgtWndHandle = FindWindow(anyPopupClass, "Downloading to target system");
                        if (osDldTgtWndHandle == IntPtr.Zero)
                        {
                            osDldTgtWndHandle = FindWindow(anyPopupClass, null);
                        }
                    }

                    if (osDldTgtWndHandle != IntPtr.Zero)
                    {
                        SetForegroundWindow(osDldTgtWndHandle);
                        IntPtr DlButtonHandle = FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "OK");
                        if (DlButtonHandle != IntPtr.Zero)
                        {
                            SendMessage(DlButtonHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                            System.Threading.Thread.Sleep(500); //important to wait a bit

                            IntPtr deactivateRTPopup = FindWindow(anyPopupClass, "Target system");
                            if (deactivateRTPopup != IntPtr.Zero)
                            {
                                SetForegroundWindow(deactivateRTPopup);
                                System.Threading.Thread.Sleep(500); //important to wait a bit
                                IntPtr YesButton = FindWindowEx(deactivateRTPopup, IntPtr.Zero, "Button", "&Yes");
                                if (YesButton != IntPtr.Zero)
                                {
                                    SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                }
                                else
                                {
                                    logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : The Ok Button was not found in the deactivation window!");
                                    logFile.Flush();
                                    MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                }
                            }
                            else
                            {
                                logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Did not find target system popup! (runtime was not active on this client)");
                                logFile.Flush();
                                //MessageBox.Show(new Form { TopMost = true }, "Did not find target system popup!"); //careful to focus on it
                            }

                            //call remote desktop open here
                            if (rdpCheckBox.Checked == true)
                            {
                                //certificate needs to be generated here
                                System.Threading.Thread.Sleep(10000); //important to wait a bit
                                OpenRemoteSession(myIp, unTextBox.Text, passTextBox.Text);
                            }

                            //if Canceled by the user in LOAD.LOG , assume that RT station not obtainable //read load.log here to find canceled by user
                            SetForegroundWindow(osDldTgtWndHandle);
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

                            SleepUntilDownloadFeedback(clientIndex + 1);

                            IntPtr dldingTgtHandle = FindWindow(anyPopupClass, "Downloading to target system");
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
                                        SetForegroundWindow(osDldTgtWndHandle);
                                        System.Threading.Thread.Sleep(500);
                                        IntPtr YesButton = FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "Ok");
                                        if (YesButton != IntPtr.Zero)
                                        {
                                            SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                        }
                                        else
                                        {
                                            logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : The Ok Button was not found in the downloading window!");
                                            logFile.Flush();
                                            MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                        }
                                        logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Error on download to client " + clientIndex + 1);
                                        logFile.Flush();
                                        continue;
                                    }

                                    if (dldingTgtText.Where(x => x.Contains("Canceled:")).Count() > 0)
                                    {
                                        logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Client " + clientIndex + " download canceled - RT station not obtainable");
                                        logFile.Flush();
                                        MessageBox.Show(new Form { TopMost = true }, " Client " + clientIndex + " download canceled - RT station not obtainable - will continue to next client download");
                                        continue;
                                    }

                                    System.Threading.Thread.Sleep(1000);

                                } while (dldingTgtText.Where(x => x.Contains("Download to target system was completed successfully")).Count() == 0);

                                ClickButtonUsingMessage(dldingTgtHandle, "OK", "Download to target system was completed successfully", logFile);
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
                        logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() +
                            " : Finished download process for machine " + checkedListBox1.CheckedItems[i].ToString() +
                            " in " + secElapsed.ToString() + " seconds");
                        logFile.Flush();
                        int currentElapsed = Convert.ToInt32(label9.Text.Split(Convert.ToChar("~"))[1]);
                        label9.Text = "Remaining [s]: ~" + (currentElapsed - 100).ToString();
                        label9.Refresh();
                        progressBar1.Refresh();
                        statusLabel.Refresh();
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Could not focus on download popup!"); //careful to focus on it
                    }
                }
                logFile.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : Closing logfile");
                logFile.Flush();
                logFile.Close();
                MessageBox.Show(new Form { TopMost = true }, "The NCM download process has been finished!"); //careful to focus on it
            }
        }

        private void ClickButtonUsingMessage(IntPtr windowHandle, string buttonText, string windowText, StreamWriter log)
        {
            bool success = true;
            do
            {
                try
                {
                    SetForegroundWindow(windowHandle);
                    IntPtr OkButton = FindWindowEx(windowHandle, IntPtr.Zero, "Button", buttonText);
                    if (OkButton != IntPtr.Zero)
                    {
                        SendMessage(OkButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, OkButton);
                        var checkText = ExtractWindowTextByHandle(windowHandle);
                        if (checkText.Where(x => x.Contains(windowText)).Count() == 0)
                        {
                            success = true;
                        }
                    }
                    else
                    {
                        log.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC : The " + buttonText + " button was not found in the downloading window to confirm finish!");
                        log.Flush();
                        //MessageBox.Show(new Form { TopMost = true }, "The OK Button was not found in the downloading to target system window!"); //careful to focus on it
                        success = false;
                    }
                }
                catch (Exception exc)
                {
                    log.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + exc.Message);
                    log.Flush();
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
                myRdp = FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
                if (myRdp != IntPtr.Zero)
                {
                    // Set the window's position.
                    int width = int.Parse(widthBox.Text);
                    int height = int.Parse(heightBox.Text);
                    int x = int.Parse(leftBox.Text);
                    int y = int.Parse(topBox.Text);
                    SetWindowPos(myRdp, IntPtr.Zero, x, y, width, height, 0);
                }
                else
                {
                    System.Threading.Thread.Sleep(1000);
                }
            } while (myRdp == IntPtr.Zero);
        }

        private void CloseRemoteSession(string ip)
        {
            IntPtr myRdp = FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
            if (myRdp != IntPtr.Zero)
            {
                SendMessage(myRdp, (int)WindowsMessages.WM_CLOSE, (int)IntPtr.Zero, IntPtr.Zero);
                System.Threading.Thread.Sleep(1000);
            }
        }


        private void SendKeyHandled(IntPtr windowHandle, string key, StreamWriter log)
        {
            bool success;
            do
            {
                try
                {
                    SetForegroundWindow(windowHandle);
                    SendKeys.Send(key);
                    success = true;
                }
                catch (Exception e)
                {
                    log.WriteLine(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + e.Message);
                    log.Flush();
                    success = false;
                    System.Threading.Thread.Sleep(10);
                }
            } while (success == false);
        }

        private void NavigateToIndex(int index, IntPtr ncm, StreamWriter log)
        {
            for (int i = 0; i < index; i++)
            {
                SendKeyHandled(ncm, "{DOWN}", log);
            }
        }

        private void ResetExpansions(IntPtr ncm, StreamWriter log)
        {
            int num = int.Parse(numClTextBox.Text) * 3;
            for (int i = 0; i < num; i++) //go up
            {
                SendKeyHandled(ncm, "{UP}", log);
            }
            for (int i = 0; i < num; i++) //expand all
            {
                SendKeyHandled(ncm, "{DOWN}", log);
                for (int j = 0; j < 6; j++)
                {
                    SendKeyHandled(ncm, "{RIGHT}", log);
                }
            }
            for (int i = 0; i < num; i++) //back to compact
            {
                for (int j = 0; j < 4; j++)
                {
                    SendKeyHandled(ncm, "{LEFT}", log);
                }
                SendKeyHandled(ncm, "{UP}", log);
            }
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{DOWN}", log); //go to first dev station or whatever in the list
        }

        private void DownloadToCurrentIndex(int index, IntPtr ncm, StreamWriter log)
        {
            NavigateToIndex(index, ncm, log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "{RIGHT}", log);
            SendKeyHandled(ncm, "^(l)", log);
            System.Threading.Thread.Sleep(500);
        }

        private void ReturnToFirstClient(IntPtr ncm, StreamWriter log)
        {
            for (int i = 0; i < 10; i++)
            {
                SendKeyHandled(ncm, "{LEFT}", log);
            }
            SendKeyHandled(ncm, "{RIGHT}", log);

            for (int i = 0; i < Int32.Parse(firstClientIndexBox.Text); i++)
            {
                SendKeyHandled(ncm, "{DOWN}", log);
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
            this.rdpBox1.SuspendLayout();
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
            this.progressBar1.Size = new System.Drawing.Size(194, 13);
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
            this.label9.Size = new System.Drawing.Size(90, 13);
            this.label9.TabIndex = 24;
            this.label9.Text = "Remaining [s]: ~0";
            // 
            // NCMForm
            // 
            this.ClientSize = new System.Drawing.Size(324, 421);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NCMForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "NCM Manager Automation";
            this.rdpBox1.ResumeLayout(false);
            this.rdpBox1.PerformLayout();
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
            label9.Text = "Remaining [s]: ~" + (checkedListBox1.CheckedIndices.Count * 100).ToString();
            label9.Refresh();
        }
    }
}