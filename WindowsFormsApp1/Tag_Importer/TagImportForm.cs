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
using FileDialog;
using InteroperabilityFunctions;
using PInvoke;

namespace Tag_Importer
{
    public partial class TagImportForm : Form
    {
        public TagImportForm()
        {
            InitializeComponent();
        }

        private void OpenTagMgmtMenu(IntPtr tagMgmt)
        {
            PInvokeLibrary.SetForegroundWindow(tagMgmt);

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
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
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
            Int32 titleSize = PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, IntPtr.Zero).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

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
            //FileDialog.Program.GetFilesInDialog();
            ImportTags();
        }

        private void ClickOnExpandOrRevert(IntPtr handle, int x, int y)
        {
            PInvokeLibrary.SetForegroundWindow(handle);

            MouseOperations.SetCursorPosition(x, y); //have to use the found minus/plus coordinates here
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
        }

        private void ImportTags()
        {
            IntPtr tag = PInvokeLibrary.FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");
            IntPtr ccAx = PInvokeLibrary.FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);
            IntPtr navBar = PInvokeLibrary.FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);
            IntPtr treeView = PInvokeLibrary.FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);
            IntPtr trHandle = PInvokeLibrary.FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", "");

            var trRect = new RECT();
            PInvokeLibrary.SetForegroundWindow(trHandle);
            PInvokeLibrary.GetWindowRect(trHandle, out trRect);
            System.Threading.Thread.Sleep(500);

            #region imageRecognitionAndAction
            Bitmap img = MyFunctions.GetPngByHandle(trHandle);
            Bitmap minus = (Bitmap)Bitmap.FromFile("minus.png");
            Bitmap plus = (Bitmap)Bitmap.FromFile("plus.png");

            List<Point> searchMinus = MyFunctions.FindBitmapsEntry(img, minus);
            List<Point> searchPlus = MyFunctions.FindBitmapsEntry(img, plus);

            //click on minus from image
            ClickOnExpandOrRevert(tag, trRect.left + searchMinus[0].X, trRect.top + searchMinus[0].Y); //have to use the found minus/plus coordinates here

            //do scrolling, same manner

            minus.Dispose();
            plus.Dispose();
            img.Dispose();
            #endregion

            OpenTagMgmtMenu(tag);

            IntPtr importPopup = PInvokeLibrary.FindWindow("#32770", "Import");
            do
            {
                try { importPopup = PInvokeLibrary.FindWindow("#32770", "Import"); }
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

            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");
            SendKeyHandled(importPopup, "{TAB}");

            //var len = SendMessage(importPopup, (uint)WindowsMessages.WM_GETTEXTLENGTH, 0, null);
            //var sb = new StringBuilder(len + 1);
            //SendMessage(importPopup, (uint)WindowsMessages.WM_GETTEXT, sb.Capacity, sb);

            IntPtr fileListParent = PInvokeLibrary.FindWindowEx(importPopup, IntPtr.Zero, "DUIViewWndClassName", null);
            IntPtr fileList = PInvokeLibrary.FindWindowEx(fileListParent, IntPtr.Zero, "DirectUIHWND", null);


            IntPtr addressBar = GetChildBySubstring("Address:", importPopup);

            Console.WriteLine("Done");
        }
    }
}
