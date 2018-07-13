using log4net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xltrail.Client
{
    public class Git
    {
        static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string path;

        public Git(string path)
        {
            this.path = path;
        }


        private void Execute(string command)
        {
            var gitInfo = new ProcessStartInfo();
            gitInfo.CreateNoWindow = true;
            gitInfo.RedirectStandardError = true;
            gitInfo.RedirectStandardOutput = true;
            gitInfo.UseShellExecute = false;
            gitInfo.FileName = "git.exe";

            Process gitProcess = new Process();
            gitInfo.Arguments = command;
            gitInfo.WorkingDirectory = path;

            logger.DebugFormat("Execute command: {0}", path);
            gitProcess.StartInfo = gitInfo;
            gitProcess.Start();

            var stderr = gitProcess.StandardError.ReadToEnd();
            var stdout = gitProcess.StandardOutput.ReadToEnd();
            gitProcess.WaitForExit();
            logger.DebugFormat("ExitCode : {0}", gitProcess.ExitCode);
            gitProcess.Close();

            logger.DebugFormat("Stdout: {0}", stdout);
            logger.DebugFormat("Stderr: {0}", stderr);

        }

        public void Commit(string message)
        {
            Execute("commit -m \"" + message + "\"");
        }

        public void Pull(string branch)
        {
            Execute("pull origin " + branch + " -X ours");
        }

        public void Fetch()
        {
            Execute("fetch origin");
        }

        public void Push(string branch)
        {
            Execute("push origin " + branch);
        }
    }
}
