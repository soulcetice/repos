using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace EnablePSRemoting
{
    class Program
    {
        static void Main(string[] args)
        {
            EnablePSRemoting();
        }

        private static void EnablePSRemoting()
        {
            //run ps command here
            //or run ps file
            //System.Diagnostics.Process.Start(Path.Combine(Application.StartupPath, "StartTrustedHosts.ps1"));
            Runspace runSpace = RunspaceFactory.CreateRunspace();
            runSpace.Open();
            Pipeline pipeline = runSpace.CreatePipeline();

            Command invokeScript = new Command("Invoke-Command");
            RunspaceInvoke invoke = new RunspaceInvoke();

            ScriptBlock sb = invoke.Invoke("{Enable-PSRemoting}")[0].BaseObject as ScriptBlock;
            invokeScript.Parameters.Add("ScriptBlock", sb);

            pipeline.Commands.Add(invokeScript);
            Collection<PSObject> output = pipeline.Invoke();
            foreach (PSObject obj in output)
            {
                //LogToFile(obj.ToString());
            }
            //LogToFile("Enabled PSremoting locally");
        }
    }
}
