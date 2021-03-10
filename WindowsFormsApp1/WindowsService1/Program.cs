using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration.Install;
using System.Reflection;

namespace WindowsService1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);
        }

        //static void Main(string[] args)
        //{
        //    if (Environment.UserInteractive)
        //    {
        //        var parameter = string.Concat(args);
        //        switch (parameter)
        //        {
        //            case "--install":
        //                ManagedInstallerClass.InstallHelper(new[] { Assembly.GetExecutingAssembly().Location });
        //                break;
        //            case "--uninstall":
        //                ManagedInstallerClass.InstallHelper(new[] { "/u", Assembly.GetExecutingAssembly().Location });
        //                break;
        //        }
        //    }
        //    else
        //    {
        //        ServiceBase[] servicesToRun = {
        //    new Service1()
        //};
        //        ServiceBase.Run(servicesToRun);
        //    }
        //}
    }
}
