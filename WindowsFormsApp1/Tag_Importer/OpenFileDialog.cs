using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using System.Security.Principal;
using System.Diagnostics;

namespace FileDialog
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, int wParam, StringBuilder lParam);

        //  [DllImport("user32.dll", CharSet = CharSet.Auto)]
        //  public static extern int SendMessage(int hWnd, int msg, int wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetDlgItem(IntPtr hwnd, int childID);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(HandleRef hwnd, int wMsg, int wParam, String s);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern String SendMessage(HandleRef hwnd, uint WM_GETTEXT);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(int hWnd, int msg, int wParam, IntPtr lParam);
        // to get file size import
        [DllImport("kernel32.dll")]
        static extern bool GetFileSizeEx(IntPtr hFile, out long lpFileSize);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CreateFile(
     [MarshalAs(UnmanagedType.LPTStr)] string filename,
     [MarshalAs(UnmanagedType.U4)] FileAccess access,
     [MarshalAs(UnmanagedType.U4)] FileShare share,
     IntPtr securityAttributes, // optional SECURITY_ATTRIBUTES struct or IntPtr.Zero
     [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
     [MarshalAs(UnmanagedType.U4)] FileAttributes flagsAndAttributes,
     IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]

        public struct WIN32_FIND_DATA
        {
            public int dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public int nFileSizeHigh;
            public int nFileSizeLow;
            public int dwReserved0;
            public int dwReserved1;
            public string cFileName; //mite need marshalling, TCHAR size = MAX_PATH???
            public string cAlternateFileName; //mite need marshalling, TCHAR size = 14
        }
        public struct WIN32_FIND_DATA1
        {
            public int dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public int nFileSizeHigh;
            public int nFileSizeLow;
            public int dwReserved0;
            public int dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindFirstFile(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

        [DllImport("kernel32.dll")]
        static extern IntPtr FindFirstFile(IntPtr lpfilename, ref WIN32_FIND_DATA findfiledata);

        [DllImport("kernel32.dll")]
        static extern IntPtr FindClose(IntPtr pff);

        [DllImport("User32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        public static Process[] myProcess = Process.GetProcessesByName("program name here");


        const uint WM_GETTEXT = 0x0D;
        const uint WM_GETTEXTLENGTH = 0X0E;
        const int BN_CLICKED = 245;
        private const int WM_SETTEXT = 0x000C;

        public static void GetFilesInDialog()
        {
            IntPtr hWnd = FindWindow(null, "Open");

            if (hWnd != IntPtr.Zero)
            {
                Console.WriteLine("Open File Dialog is open");

                IntPtr hwndButton = FindWindowEx(hWnd, IntPtr.Zero, "Button", "&Open");
                Console.WriteLine("The handle of the Open button is " + hwndButton);

                IntPtr FileDialogHandle = FindWindow(null, "Open");
                IntPtr iptrHWndControl = GetDlgItem(FileDialogHandle, 1148);
                HandleRef hrefHWndTarget = new HandleRef(null, iptrHWndControl);
                //SendMessage(hrefHWndTarget, WM_SETTEXT, 0, "your file path");

                IntPtr opnButton = FindWindowEx(FileDialogHandle, IntPtr.Zero, "Open", null);

                SendMessage((int)opnButton, BN_CLICKED, 0, IntPtr.Zero);

                int len = (int)SendMessage(hrefHWndTarget, WM_GETTEXTLENGTH, 0, null);
                var sb = new StringBuilder(len + 1);

                SendMessage(hrefHWndTarget, WM_GETTEXT, sb.Capacity, sb);
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
            GetFileSizeEx(handle, out fileSize);
            CloseHandle(handle);
            return (uint)fileSize;
        }
    }
}