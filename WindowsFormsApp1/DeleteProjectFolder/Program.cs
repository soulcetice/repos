using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using toolsforimpersonations;

namespace DeleteProjectFolder
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var exeName = System.AppDomain.CurrentDomain.FriendlyName;
            currentPath = currentPath.Substring(0, currentPath.Length - exeName.Length);
            var projectPathFile = currentPath + @"DeleteProjectFolderSettings.txt";

            List<string> path = GetData(projectPathFile);

            var msg = "Started deleting project folder on ...";

            //UserImpersonation impersonator = new UserImpersonation();
            //impersonator.impersonateUser("Administrator", "", "!sysadmin"); //No Domain is required

            try
            {
                if (new DirectoryInfo(path[0]).Exists)
                {
                    System.IO.DirectoryInfo di = new DirectoryInfo(path[0]);

                    foreach (FileInfo file in di.GetFiles())
                    {
                        if (!file.Extension.ToLower().Contains("mcp"))
                        {
                            file.Delete();
                        }
                    }
                    foreach (DirectoryInfo dir in di.GetDirectories())
                    {
                        if (dir.Name.ToLower() != "sdib_tcm_clt")
                        {
                            dir.Delete(true);
                        }
                    }
                }
                else
                {
                    msg = "Tried to delete, folder did not exist";
                }
            }
            catch (Exception exc)
            {
                msg = exc.Message;
            }
            //impersonator.undoimpersonateUser();

            Console.WriteLine(path);
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

namespace toolsforimpersonations
{
    public class Impersonator
    {
        #region "Consts"

        public const int LOGON32_LOGON_INTERACTIVE = 2;

        public const int LOGON32_PROVIDER_DEFAULT = 0;
        #endregion

        #region "External API"
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool RevertToSelf();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int CloseHandle(IntPtr hObject);

        #endregion

        #region "Methods"

        //Public Sub PerformImpersonatedTask(ByVal username As String, ByVal domain As String, ByVal password As String, ByVal logonType As Integer, ByVal logonProvider As Integer, ByVal methodToPerform As Action)
        public void PerformImpersonatedTask(string username, string domain, string password, int logonType, int logonProvider, Action methodToPerform)
        {
            IntPtr token = IntPtr.Zero;
            if (RevertToSelf())
            {
                if (LogonUser(username, domain, password, logonType, logonProvider, ref token) != 0)
                {
                    dynamic identity = new WindowsIdentity(token);
                    dynamic impersonationContext = identity.Impersonate();
                    if (impersonationContext != null)
                    {
                        methodToPerform.Invoke();
                        impersonationContext.Undo();
                    }
                    // do logging
                }
                else
                {
                }
            }
            if (token != IntPtr.Zero)
            {
                CloseHandle(token);
            }
        }

        #endregion
    }

    public class UserImpersonation
    {
        const int LOGON32_LOGON_INTERACTIVE = 2;
        const int LOGON32_LOGON_NETWORK = 3;
        const int LOGON32_LOGON_BATCH = 4;
        const int LOGON32_LOGON_SERVICE = 5;
        const int LOGON32_LOGON_UNLOCK = 7;
        const int LOGON32_LOGON_NETWORK_CLEARTEXT = 8;
        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_PROVIDER_WINNT35 = 1;
        const int LOGON32_PROVIDER_WINNT40 = 2;
        const int LOGON32_PROVIDER_WINNT50 = 3;

        WindowsImpersonationContext impersonationContext;
        [DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int LogonUserA(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]

        public static extern int DuplicateToken(IntPtr ExistingTokenHandle, int ImpersonationLevel, ref IntPtr DuplicateTokenHandle);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]

        public static extern long RevertToSelf();
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern long CloseHandle(IntPtr handle);

        public bool impersonateUser(string userName, string domain, string password)
        {
            return impersonateValidUser(userName, domain, password);
        }

        public void undoimpersonateUser()
        {
            undoImpersonation();
        }

        private bool impersonateValidUser(string userName, string domain, string password)
        {
            bool functionReturnValue = false;

            WindowsIdentity tempWindowsIdentity = null;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;
            functionReturnValue = false;

            //if (RevertToSelf()) {
            if (LogonUserA(userName, domain, password, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_WINNT50, ref token) != 0)
            {
                if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                {
                    tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                    impersonationContext = tempWindowsIdentity.Impersonate();
                    if ((impersonationContext != null))
                    {
                        functionReturnValue = true;
                    }
                }
            }
            //}
            if (!tokenDuplicate.Equals(IntPtr.Zero))
            {
                CloseHandle(tokenDuplicate);
            }
            if (!token.Equals(IntPtr.Zero))
            {
                CloseHandle(token);
            }
            return functionReturnValue;
        }

        private void undoImpersonation()
        {
            impersonationContext.Undo();
        }
    }
}
