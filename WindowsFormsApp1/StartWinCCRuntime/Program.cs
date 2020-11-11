using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

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
            var currentPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var exeName = System.AppDomain.CurrentDomain.FriendlyName;
            currentPath = currentPath.Substring(0, currentPath.Length - exeName.Length);
            var dataFile = currentPath + @"StartWinCCRuntimeSettings.txt";

            List<string> data = GetData(dataFile);
            var myPath = data[0];
            var winccex = data[1];
            var pdlrt = data[2];


            Process[] processlist = Process.GetProcessesByName("WinCCExplorer");
            Process.Start(winccex, myPath);
            //WinCC Explorer - C:\Users\admin\Downloads\Projects\ELV-HFM\wincproj\ELVAL_HFM_CLT\ELVAL_HFM_CLT.MCP
            var path = "WinCC Explorer - " + myPath;
            System.Diagnostics.Debug.WriteLine(path);
            var flag = false;
            do
            {
                processlist = Process.GetProcessesByName("WinCCExplorer");
                foreach (var proc in processlist)
                {
                    if (proc.ProcessName == "WinCCExplorer")
                    {
                        if (proc.MainWindowTitle == path)
                        {
                            flag = true;
                        }
                    }
                }
                System.Threading.Thread.Sleep(1000);
            } while (flag == false);

            Process.Start(pdlrt);
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
