using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowHandle;

namespace Tag_Importer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            IntPtr tag = FindWindow("WinCC ConfigurationStudio MainWindow", "Tag Management - WinCC Configuration Studio");

            IntPtr ccAx = FindWindowEx(tag, IntPtr.Zero, "CCAxControlContainerWindow", null);

            IntPtr navBar = FindWindowEx(ccAx, IntPtr.Zero, "WinCC NavigationBarControl Window", null);

            IntPtr treeView = FindWindowEx(navBar, IntPtr.Zero, "WinCC ConfigurationStudio NavigationBarTreeView", null);

            IntPtr trHandle = FindWindowEx(treeView, IntPtr.Zero, "SysTreeView32", null);

            Console.Write("Done");
        }

        #region ImportDlls
        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //get objects in window ?
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr handleParent, IntPtr handleChild, string className, string WindowName);
        #endregion
    }
}
