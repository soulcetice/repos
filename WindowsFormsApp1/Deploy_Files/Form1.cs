using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsUtilities;

namespace Deploy_Files
{
    public partial class Form1 : Form
    {
        public Form1()
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

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);
        //get objects in window ?
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr handleParent, IntPtr handleChild, string className, string WindowName);
        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        private void button2_Click(object sender, EventArgs e)
        {
            //    IntPtr myRdp = FindWindow("TscShellContainerClass", "10.127.2.166 - Remote Desktop Connection");
            //    SendMessage(myRdp, (int)WindowsMessages.WM_CLOSE, (int)IntPtr.Zero, IntPtr.Zero);
            //    IntPtr rdpPopup = FindWindow("#32770", "Remote Desktop Connection");
            //    //SetForegroundWindow(rdpPopup);
            //    //System.Threading.Thread.Sleep(50);
            //    //SendKeys.Send("{ENTER}");

            //    //IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", null);
            //    SendMessage(IntPtr.Zero, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);

            IntPtr rdpPopup = FindWindow("#32770", "Download OS");
            SetForegroundWindow(rdpPopup);
            //System.Threading.Thread.Sleep(50);
            //SendKeys.Send("{ENTER}");

            //IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", null);
            IntPtr btnHandle = FindWindowEx(rdpPopup, IntPtr.Zero, "Button", "OK");
            SendMessage(btnHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
        }
    }
}
