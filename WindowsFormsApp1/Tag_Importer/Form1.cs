using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowHandle;
using WindowsUtilities;

namespace Tag_Importer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void OpenTagMgmtMenu(IntPtr tagMgmt)
        {
            SetForegroundWindow(tagMgmt);

            SendKeyHandled(tagMgmt, "(%)");
            SendKeyHandled(tagMgmt, "{RIGHT}");
            SendKeyHandled(tagMgmt, "{ENTER}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{DOWN}");
            //SendKeyHandled(tagMgmt, "{DOWN}");
            SendKeyHandled(tagMgmt, "{ENTER}");
        }

        private void SendKeyHandled(IntPtr windowHandle, string key/*, StreamWriter log*/)
        {
            bool success;
            do
            {
                try
                {
                    SetForegroundWindow(windowHandle);
                    SendKeys.SendWait(key);
                    success = true;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    //log.WriteLine(e.Message);
                    //log.Flush();
                    success = false;
                }
            } while (success == false);
        }

        #region ImportDlls
        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //get objects in window ?
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr handleParent, IntPtr handleChild, string className, string WindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);
        [DllImport("kernel32.dll")]
        static extern int GetProcessId(IntPtr handle);

        [DllImport("kernel32.dll")]
        public static extern bool ReadProcessMemory(int hProcess, int lpBaseAddress, byte[] lpBuffer, int dwSize, ref int lpNumberOfBytesRead);
        #endregion

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

        public string GetControlText(IntPtr hWnd)
        {

            // Get the size of the string required to hold the window title (including trailing null.) 
            Int32 titleSize = SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, IntPtr.Zero).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

            return title.ToString();
        }

        public IntPtr GetChildBySubstring(string str, IntPtr handle)
        {
            var extractedText = new List<string>();
            List<IntPtr> childObjects = new WindowHandleInfo(handle).GetAllChildHandles();
            for (int i = 0; i < childObjects.Count; i++)
            {
                var text = GetControlText(childObjects[i]);
                if (text.Contains(str))
                {
                    return childObjects[i];
                }
            }
            return IntPtr.Zero;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IntPtr tag = FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");

            IntPtr ccAx = FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);

            IntPtr navBar = FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);

            IntPtr treeView = FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);

            IntPtr trHandle = FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", null);

            OpenTagMgmtMenu(tag);

            IntPtr importPopup = FindWindow("#32770", "Import");
            do
            {
                try { importPopup = FindWindow("#32770", "Import"); }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.Message);
                }
                System.Threading.Thread.Sleep(100);
            } while (importPopup == IntPtr.Zero);

            //changing address...
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{ENTER}");
            string path = textBox1.Text;
            foreach (char c in path)
                SendKeys.SendWait(c.ToString());
            SendKeyHandled(importPopup, "{ENTER}");


            IntPtr fileListParent = FindWindowEx(importPopup, IntPtr.Zero, "DUIViewWndClassName", null);
            IntPtr fileList = FindWindowEx(fileListParent, IntPtr.Zero, "DirectUIHWND", null);


            //Process myProcess = Process.GetProcesses().Single(p => p.Id != 0 && p.Handle == importPopup);

            //var proc = Process.GetProcesses().FirstOrDefault(c => c.ProcessName == "CCConfigStudio_x64").MainModule.

            Class1.ReadProcessData("notepad");

            IntPtr addressBar = GetChildBySubstring("Address:", importPopup);

            Console.WriteLine("Done");
        }
    }
}
