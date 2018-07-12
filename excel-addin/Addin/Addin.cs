using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.IO;
using Excel=Microsoft.Office.Interop.Excel;
using log4net;
using System.Threading;
using System.Diagnostics;
using System.Linq;

namespace xltrail.Client
{
    public static class Addin
    {
        [ComVisible(true)]
        public class RibbonController : ExcelRibbon, IExcelAddIn
        {
            static Excel.Application xlApp;
            static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            IRibbonUI ribbon;


            public void AutoOpen()
            {
                //ogger.Setup();
                logger.Info("Initialise Addin");
                xlApp = (Excel.Application)ExcelDnaUtil.Application;
                //xlApp.WorkbookActivate += XlApp_WorkbookActivate;
                xlApp.WorkbookOpen += XlApp_WorkbookOpen;
                xlApp.WorkbookAfterSave += XlApp_WorkbookAfterSave;
            }

            private void XlApp_WorkbookOpen(Excel.Workbook workbook)
            {
                xlApp.ScreenUpdating = true;
                var path = workbook.FullName;

                var directory = Path.GetDirectoryName(path);
                if (!LibGit2Sharp.Repository.IsValid(directory))
                    return;

                //workbook path relative to repository
                var gitPath = path.Substring(directory.Length + 1).Replace("\\", "/");

                //libgit2 repo
                var repository = new LibGit2Sharp.Repository(directory);
                var branch = repository.Head.FriendlyName;

                //git command line wrapper
                var git = new Git(directory);


                //if workbook modified, no need to fetch from origin as we cannot do anything anyway
                var workbookStatus = repository.RetrieveStatus(gitPath);
                if (workbookStatus == LibGit2Sharp.FileStatus.ModifiedInWorkdir)
                    return;

                //fetch from origin (if exists)
                var origin = repository.Network.Remotes.Where(x => x.Name == "origin").FirstOrDefault();
                if (origin != null)
                {
                    var url = origin.PushUrl;
                    xlApp.StatusBar = "Fetching from " + url;
                    git.Fetch();
                    xlApp.StatusBar = null;
                    var remote = repository.Branches.Where(b => b.IsRemote).Where(b => b.FriendlyName == "origin/" + branch).FirstOrDefault();
                    var local = repository.Branches[branch];


                    //check if local != remote
                    if (local.Tip.Id != remote.Tip.Id)
                    {
                        var localWorkbook = local[gitPath];
                        var remoteWorkbook = remote[gitPath];

                        var localFileSha = ((LibGit2Sharp.Blob)localWorkbook.Target).Sha;
                        var remoteFileSha = ((LibGit2Sharp.Blob)remoteWorkbook.Target).Sha;

                        xlApp.ScreenUpdating = false;

                        //close active workbook
                        workbook.Close(false);

                        //reset local to remotr
                        repository.Reset(LibGit2Sharp.ResetMode.Hard, remote.Tip);

                        //merge remote into current branch
                        xlApp.Workbooks.Open(path);
                    }

                }
            }

            private void XlApp_WorkbookAfterSave(Excel.Workbook workbook, bool Success)
            {
                var path = workbook.FullName;
                var name = workbook.Name;

                var directory = Path.GetDirectoryName(path);
                if (!LibGit2Sharp.Repository.IsValid(directory))
                    return;

                //libgit2 repo
                var repository = new LibGit2Sharp.Repository(directory);
                var branch = repository.Head.FriendlyName;

                //git command line wrapper
                var git = new Git(directory);

                //add
                LibGit2Sharp.Commands.Stage(repository, path);

                //commit
                git.Commit("Modified " + name);

                //push to origin (if exists)
                var origin = repository.Network.Remotes.Where(x => x.Name == "origin").FirstOrDefault();
                if (origin != null)
                {
                    var pushUrl = origin.PushUrl;
                    xlApp.StatusBar = "Pushing to " + pushUrl;
                    git.Pull(branch);
                    git.Push(branch);
                    xlApp.StatusBar = null;
                }
            }

            private void ShowNotification(string description)
            {
                var notification = new System.Windows.Forms.NotifyIcon()
                {
                    Visible = true,
                    Icon = System.Drawing.SystemIcons.Information,
                    // optional - BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info,
                    // optional - BalloonTipTitle = "My Title",
                    BalloonTipText = description,
                };

                // Display for 5 seconds.
                notification.ShowBalloonTip(5000);

                // This will let the balloon close after it's 5 second timeout
                // for demonstration purposes. Comment this out to see what happens
                // when dispose is called while a balloon is still visible.
                Thread.Sleep(10000);

                // The notification should be disposed when you don't need it anymore,
                // but doing so will immediately close the balloon if it's visible.
                notification.Dispose();

            }

            public void AutoClose()
            {
            }


            /*
            public void Ribbon_Load(IRibbonUI ribbon)
            {
                this.ribbon = ribbon;
            }*/

            /*            
            private void XlApp_WorkbookActivate(Excel.Workbook Wb)
            {
                var path = xlApp.ActiveWorkbook.FullName;
                if (!path.Contains(StagingPath))
                {
                    activeWorkbookBranch = null;
                }
                else
                {
                    activeWorkbookBranch = repositories.GetWorkbookVersionFromPath(GetWorkbookPath(path));
                }
                ribbon.Invalidate();
            }*/
        }
    }
}
