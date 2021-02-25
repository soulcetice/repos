using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace StartWinCCRuntime
{
    class Program
    {
        static void Main(string[] args)
        {
            Actions();
        }

        static void Actions()
        {
            var c = new DirectoryInfo(@"C:\Project\");
            var d = new DirectoryInfo(@"D:\Project\");
            DirectoryInfo cltDir = null;
            bool dirFlag = false;
            while (dirFlag == false)
            {
                if (c.Exists)
                {
                    FindClientFolder(c, ref cltDir, ref dirFlag);
                }
                else if (d.Exists)
                {
                    FindClientFolder(d, ref cltDir, ref dirFlag);
                }
            }

            if (cltDir == null)
                return;

            var myPath = "";
            foreach (var file in cltDir.GetFiles())
            {
                if (file.Extension.ToLower() == ".mcp")
                {
                    myPath = file.FullName;
                }
            }

            if (myPath == null)
                return;

            var winccex = @"C:\Program Files (x86)\Siemens\WinCC\bin\WinCCExplorer.exe";
            var pdlrt = @"C:\Program Files (x86)\Siemens\WinCC\bin\PdlRt.exe";

            Debug.Print(myPath);

            Process.Start(new ProcessStartInfo
            {
                Arguments = "/C " + @"""" + winccex + @""" " + myPath,
                FileName = "cmd",
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = @"C:\Program Files (x86)\Siemens\WinCC\bin\" //<---very important
            });

            //double wait = 0;
            //bool flag = false;
            //IntPtr explorerWindow;
            //do
            //{
            //    explorerWindow = WndSearcher.SearchForWindow("WinCCExplorerFrameWndClass", "WinCC Explorer -");
            //    if (wait > 15000 || explorerWindow != IntPtr.Zero)
            //        flag = true;
            //    System.Threading.Thread.Sleep(1000);
            //    wait += 1000;
            //} while (flag == false);

            System.Threading.Thread.Sleep(15000);

            Process.Start(new ProcessStartInfo
            {
                Arguments = "/C " + @"""" + pdlrt + @"""",
                FileName = "cmd",
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = @"C:\Program Files (x86)\Siemens\WinCC\bin\" //<---very important
            });
            //impersonator.undoimpersonateUser();
        }

        private static void FindClientFolder(DirectoryInfo c, ref DirectoryInfo cltDir, ref bool dirFlag)
        {
            var dirs = c.GetDirectories();
            foreach (var dir in dirs)
            {
                if (dir.Name.ToLower().EndsWith("_clt"))
                {
                    cltDir = dir;
                    dirFlag = true;
                }
            }
        }

        private static List<string> GetData(string settingsFile)
        {
            var data = new List<string>();
            using (var fileReader = new StreamReader(settingsFile, true))
            {
                while (!fileReader.EndOfStream)
                    data.Add(fileReader.ReadLine());
                fileReader.Close();
            }
            return data;
        }
    }
}


public class WndSearcher
{
    public static IntPtr SearchForWindow(string wndclass, string title)
    {
        SearchData sd = new SearchData { Wndclass = wndclass, Title = title };
        EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
        return sd.hWnd;
    }

    public static bool EnumProc(IntPtr hWnd, ref SearchData data)
    {
        // Check classname and title
        // This is different from FindWindow() in that the code below allows partial matches
        StringBuilder sb = new StringBuilder(1024);
        GetClassName(hWnd, sb, sb.Capacity);
        if (sb.ToString().StartsWith(data.Wndclass))
        {
            sb = new StringBuilder(1024);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().StartsWith(data.Title))
            {
                data.hWnd = hWnd;
                return false;    // Found the wnd, halt enumeration
            }
        }
        return true;
    }

    public class SearchData
    {
        // You can put any dicks or Doms in here...
        public string Wndclass;
        public string Title;
        public IntPtr hWnd;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref SearchData data);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
}
