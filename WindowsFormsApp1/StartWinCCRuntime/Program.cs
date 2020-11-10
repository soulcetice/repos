using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace StartWinCCRuntime
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var exeName = System.AppDomain.CurrentDomain.FriendlyName;
            currentPath = currentPath.Substring(0, currentPath.Length - exeName.Length);
            var dataFile = currentPath + @"StartWinCCRuntimeSettings.txt";

            List<string> data = GetData(dataFile);
            var myPath = data[0];
            var winccex = data[1];
            var pdlrt = data[2];

            Process.Start(winccex, myPath);
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
