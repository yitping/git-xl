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

            gitProcess.StartInfo = gitInfo;
            gitProcess.Start();

            string stderr = gitProcess.StandardError.ReadToEnd();
            string stdout = gitProcess.StandardOutput.ReadToEnd();

            gitProcess.WaitForExit();
            gitProcess.Close();
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
