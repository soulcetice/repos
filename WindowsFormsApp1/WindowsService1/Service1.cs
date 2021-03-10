using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using CommonInterops;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        public Service1(/*string[] args*/)
        {
            InitializeComponent();

            string eventSourceName = "MySource";
            string logName = "MyNewLog";

            //if (args.Length > 0)
            //{
            //    eventSourceName = args[0];
            //}

            //if (args.Length > 1)
            //{
            //    logName = args[1];
            //}

            var eventLog1 = new EventLog();

            if (!EventLog.SourceExists(eventSourceName))
            {
                EventLog.CreateEventSource(eventSourceName, logName);
            }

            eventLog1.Source = eventSourceName;
            eventLog1.Log = logName;
        }

        protected override void OnStart(string[] args)
        {
            //(Environment.UserName);
        }

        protected override void OnStop()
        {
        }
    }
}
