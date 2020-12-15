using System;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace StopWinCCRuntime
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();
                rt.Stop();

                System.Threading.Thread.Sleep(500);
            }
            catch (Exception exc)
            {
                string ex = exc.Message;
            }
            finally
            {
                var anyPopupClass = "#32770"; //usually any popup
                IntPtr deactivatingPopup = WndSearcher.SearchForWindow(anyPopupClass, "Deactivating -");
                do
                {
                    deactivatingPopup = WndSearcher.SearchForWindow(anyPopupClass, "Deactivating -");
                    System.Threading.Thread.Sleep(500);
                } while (deactivatingPopup != IntPtr.Zero);

                System.Threading.Thread.Sleep(5000);

                //System.Diagnostics.Process.Start(@"cscript //B //Nologo C:\Program Files (x86)\Siemens\WinCC\bin\Reset_WinCC.vbs"); 
                Process scriptProc = new Process();
                scriptProc.StartInfo.FileName = @"cscript";
                scriptProc.StartInfo.WorkingDirectory = @"C:\Program Files (x86)\Siemens\WinCC\bin\"; //<---very important 
                scriptProc.StartInfo.Arguments = "//B //Nologo Reset_WinCC.vbs";
                scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up
                scriptProc.Start();
                scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exit
                scriptProc.Close();
                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Siemens\WinCC\bin\CCCleaner.exe", "-terminate"); //not working uac
            }
        }

        public delegate bool EnumedWindow(IntPtr handleWindow, ArrayList handles);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]

        public static extern bool EnumWindows(EnumedWindow lpEnumFunc, ArrayList lParam);

        public static ArrayList GetWindows()
        {
            ArrayList windowHandles = new ArrayList();
            EnumedWindow callBackPtr = GetWindowHandle;
            EnumWindows(callBackPtr, windowHandles);

            return windowHandles;
        }

        private static bool GetWindowHandle(IntPtr windowHandle, ArrayList windowHandles)
        {
            windowHandles.Add(windowHandle);
            return true;
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