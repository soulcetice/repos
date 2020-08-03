using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsUtilities;
using Interoperability;
using System.Text;

namespace Deploy_Files
{
    public partial class DeployForm : Form
    {
        public DeployForm()
        {
            InitializeComponent();

            textBox5_TextChanged(textBox5, new System.EventArgs());

            if (File.Exists(Application.ExecutablePath + @".ini"))
            {
                List<string> data = File.ReadLines(Application.ExecutablePath + @".ini").Skip(19).Take(7).ToList();
                var ips = data[3].Split(Convert.ToChar(","));
                foreach (var item in ips)
                {
                    if (item != string.Empty) checkedListBox2.Items.Add(item);
                }

                textBox2.Text = data[4];
                textBox3.Text = data[5];
                textBox4.Text = data[0];
                textBox6.Text = data[1];
            }
        }

        public List<dynamic> fullFiles = new List<dynamic>();
        public List<dynamic> safeFiles = new List<dynamic>();

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Value = dateTimePicker2.Value.AddHours(Convert.ToDouble(textBox5.Text));
        }


        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            fullFiles.AddRange(openFileDialog1.FileNames);
            safeFiles.AddRange(openFileDialog1.SafeFileNames);
            foreach (var item in safeFiles)
            {
                checkedListBox1.Items.Add(item);
            }
            Console.WriteLine("nice");
        }

        private void SendKey(IntPtr key, IntPtr handle)
        {
            Interoperability.PInvokeLibrary.PostMessage(handle, (uint)WindowsMessages.WM_KEYDOWN, key, IntPtr.Zero);
            Interoperability.PInvokeLibrary.PostMessage(handle, (uint)WindowsMessages.WM_KEYUP, key, IntPtr.Zero);
        }

        private void SendKeyHandled(IntPtr ncm, string key)
        {
            Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
            bool success;
            do
            {
                try
                {
                    SendKeys.Send(key);
                    success = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    success = false;
                }
            } while (success == false);
        }

        private void ResetExpansions(IntPtr ncm)
        {
            System.Threading.Thread.Sleep(1000);
            int num = 12;
            for (int i = 0; i < num; i++) //go up
            {
                Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
                SendKeyHandled(ncm, "{UP}");
            }
            for (int i = 0; i < num; i++) //expand all
            {
                Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
                SendKeyHandled(ncm, "{DOWN}");
                for (int j = 0; j < 6; j++)
                {
                    Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
                    SendKeyHandled(ncm, "{RIGHT}");
                }
            }
            for (int i = 0; i < num; i++) //back to compact
            {
                for (int j = 0; j < 4; j++)
                {
                    Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
                    SendKeyHandled(ncm, "{LEFT}");
                }
                Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
                SendKeyHandled(ncm, "{UP}");
            }
            Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
            SendKeyHandled(ncm, "{RIGHT}");
            Interoperability.PInvokeLibrary.SetForegroundWindow(ncm);
            SendKeyHandled(ncm, "{DOWN}"); //go to first dev station or whatever in the list
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //tests();
            Start();
        }

        private void tests()
        {
            var ncmWndClass = "s7tgtopx"; //ncm manager main window
            var anyPopupClass = "#32770"; //usually any popup

            IntPtr ncmHandle = Interoperability.PInvokeLibrary.FindWindow(ncmWndClass, null);
            IntPtr tgtWndHandle = Interoperability.PInvokeLibrary.FindWindow(anyPopupClass, null);

            //    IntPtr myRdp = FindWindow("TscShellContainerClass", "10.127.2.166 - Remote Desktop Connection");
            //    SendMessage(myRdp, (int)WindowsMessages.WM_CLOSE, (int)IntPtr.Zero, IntPtr.Zero);
            //    IntPtr rdpPopup = FindWindow("#32770", "Remote Desktop Connection");
            //    //SetForegroundWindow(rdpPopup);
            //    //System.Threading.Thread.Sleep(50);
            //    //SendKeys.Send("{ENTER}");

            //    //IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", null);
            //    SendMessage(IntPtr.Zero, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);

            //IntPtr rdpPopup = FindWindow("#32770", "Download OS");
            //SetForegroundWindow(rdpPopup);
            ////System.Threading.Thread.Sleep(50);
            ////SendKeys.Send("{ENTER}");

            ////IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", null);
            //IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", "OK");
            //SendMessage(btnHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);


            //IntPtr dldingTgtHandle = FindWindow("#32770", "Downloading to target system");
            //SetForegroundWindow(dldingTgtHandle);
            //IntPtr OkButton = FindWindowEx(dldingTgtHandle, IntPtr.Zero, "Button", "OK");
            //SendMessage(OkButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, OkButton);

            Interoperability.PInvokeLibrary.SetForegroundWindow(ncmHandle);
            System.Threading.Thread.Sleep(500);
            ResetExpansions(ncmHandle);
        }

        //Define TreeView Flags and Messages
        private const int BN_CLICKED = 0xF5;
        private const int TV_FIRST = 0x1100;
        private const int TVGN_ROOT = 0x0;
        private const int TVGN_NEXT = 0x1;
        private const int TVGN_CHILD = 0x4;
        private const int TVGN_FIRSTVISIBLE = 0x5;
        private const int TVGN_NEXTVISIBLE = 0x6;
        private const int TVGN_CARET = 0x9;
        private const int TVM_SELECTITEM = (TV_FIRST + 11);
        private const int TVM_GETNEXTITEM = (TV_FIRST + 10);
        private const int TVM_GETITEM = (TV_FIRST + 12);

        public static void Start()
        {
            //Handle variables
            IntPtr hwnd = IntPtr.Zero;
            //int treeItem = 0;
            IntPtr hwndChild = IntPtr.Zero;
            IntPtr treeChild = IntPtr.Zero;
            var ncmWndClass = "s7tgtopx"; //ncm manager main window
            var anyPopupClass = "#32770"; //usually any popup


            hwnd = Interoperability.PInvokeLibrary.FindWindow(ncmWndClass, null);

            if (hwnd != IntPtr.Zero)
            {
                //Select TreeView Item
                var parent3 = Interoperability.PInvokeLibrary.FindWindowEx(hwnd, IntPtr.Zero, "MDIClient", null);
                var parent2 = Interoperability.PInvokeLibrary.FindWindowEx(parent3, IntPtr.Zero, "Afx:400000:b:10003:6:7fde068d", null);
                var parent1 = Interoperability.PInvokeLibrary.FindWindowEx(parent2, IntPtr.Zero, "AfxFrameOrView42", null);
                var parent = Interoperability.PInvokeLibrary.FindWindowEx(parent1, IntPtr.Zero, anyPopupClass, null);
                var treeHandle = Interoperability.PInvokeLibrary.FindWindowEx(parent, IntPtr.Zero, "SysTreeView32", null);
                treeChild = treeHandle;

                var t = new StringBuilder();
                int b = 0;
                var treeItem = Interoperability.PInvokeLibrary.SendMessage(treeChild, (int)WindowsMessages.TVM_GETCOUNT, 0, (IntPtr)b);
                //treeItem = Interoperability.PInvokeLibrary.SendMessage((int)treeChild, TVM_GETNEXTITEM, TVGN_NEXT, (IntPtr)treeItem);
                //treeItem = Interoperability.PInvokeLibrary.SendMessage((int)treeChild, TVM_GETNEXTITEM, TVGN_CHILD, (IntPtr)treeItem);
                //Interoperability.PInvokeLibrary.SendMessage((int)treeChild, TVM_SELECTITEM, TVGN_CARET, (IntPtr)treeItem);


                // ...Continue with my automation...
            }
        }//END Scan
    }
}
