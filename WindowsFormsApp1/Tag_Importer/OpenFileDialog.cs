using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using System.Security.Principal;
using System.Diagnostics;
using Tag_Importer;
using InteroperabilityFunctions;

namespace FileDialog
{
    class Program
    {

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CreateFile(
     [MarshalAs(UnmanagedType.LPTStr)] string filename,
     [MarshalAs(UnmanagedType.U4)] FileAccess access,
     [MarshalAs(UnmanagedType.U4)] FileShare share,
     IntPtr securityAttributes, // optional SECURITY_ATTRIBUTES struct or IntPtr.Zero
     [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
     [MarshalAs(UnmanagedType.U4)] FileAttributes flagsAndAttributes,
     IntPtr templateFile);


        public static Process[] myProcess = Process.GetProcessesByName("program name here");


        const uint WM_GETTEXT = 0x0D;
        const uint WM_GETTEXTLENGTH = 0X0E;
        const int BN_CLICKED = 245;
        private const int WM_SETTEXT = 0x000C;

        public static void GetFilesInDialog()
        {
            IntPtr hWnd = InteroperabilityFunctions.PInvokeLibrary.FindWindow(null, "Untitled - Notepad");

            if (hWnd != IntPtr.Zero)
            {

                //// Prepare the WINDOWPLACEMENT structure.
                //var placement = new NCMForm.WINDOWPLACEMENT();
                //placement.Length = Marshal.SizeOf(placement);
                //// Get the window's current placement.
                //NCMForm.GetWindowPlacement(hWnd, ref placement);
                //placement.ShowCmd = NCMForm.ShowWindowCommands.Maximize;
                //NCMForm.SetWindowPlacement(hWnd, ref placement);

                Console.WriteLine("Open File Dialog is open");

                IntPtr hwndButton = InteroperabilityFunctions.PInvokeLibrary.FindWindowEx(hWnd, IntPtr.Zero, "Button", "&Open");
                Console.WriteLine("The handle of the Open button is " + hwndButton);

                IntPtr FileDialogHandle = InteroperabilityFunctions.PInvokeLibrary.FindWindow(null, "Open");

                IntPtr iptrHWndControl = InteroperabilityFunctions.PInvokeLibrary.GetDlgItem(FileDialogHandle, 1148);
                HandleRef hrefHWndTarget = new HandleRef(null, iptrHWndControl);
                InteroperabilityFunctions.PInvokeLibrary.SendMessage(hrefHWndTarget, WM_SETTEXT, 0, "your file path");

                IntPtr opnButton = InteroperabilityFunctions.PInvokeLibrary.FindWindowEx(FileDialogHandle, IntPtr.Zero, "Open", null);

                InteroperabilityFunctions.PInvokeLibrary.SendMessage((int)opnButton, BN_CLICKED, 0, IntPtr.Zero);

                int len = (int)InteroperabilityFunctions.PInvokeLibrary.SendMessage(hrefHWndTarget, WM_GETTEXTLENGTH, 0, null);
                var sb = new StringBuilder(len + 1);

                InteroperabilityFunctions.PInvokeLibrary.SendMessage(hrefHWndTarget, WM_GETTEXT, sb.Capacity, sb);
                string text = sb.ToString();
                FileInfo f = new FileInfo(text);
                DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"c:\");
                FileInfo[] filesInDir = hdDirectoryInWhichToSearch.GetFiles("*" + text + "*.*");

                foreach (FileInfo foundFile in filesInDir)
                {
                    string fullName = foundFile.FullName;
                    Console.WriteLine(fullName);
                }

                var newName = DateTime.Now;

                var Username = (WindowsIdentity.GetCurrent().Name);
                var contentArray = GetFileSizeB(text);

                Console.WriteLine("The Edit box contains " + text + "\tsize:" + contentArray + "\nUser Name " + Username + "\tTime : " + newName);
            }
            else
            {
                Console.WriteLine("Open File Dialog is not open");
            }

            Console.ReadKey();
        }

        public static uint GetFileSizeB(string filename)
        {
            IntPtr handle = CreateFile(
                filename,
                FileAccess.Read,
                FileShare.Read,
                IntPtr.Zero,
                FileMode.Open,
                FileAttributes.ReadOnly,
                IntPtr.Zero);
            if (handle.ToInt32() == -1)
            {
                return 1;
            }
            long fileSize;
            InteroperabilityFunctions.PInvokeLibrary.GetFileSizeEx(handle, out fileSize);
            InteroperabilityFunctions.PInvokeLibrary.CloseHandle(handle);
            return (uint)fileSize;
        }
    }
}