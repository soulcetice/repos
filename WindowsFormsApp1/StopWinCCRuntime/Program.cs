using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StopWinCCRuntime
{
    class Program
    {
        static void Main(string[] args)
        {
            CCHMIRUNTIME.HMIRuntime rt = new CCHMIRUNTIME.HMIRuntime();
            rt.Stop();
            

            //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Siemens\WinCC\bin\CCCleaner.exe", "-terminate"); //not working uac
        }
    }
}
