using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.IO;
using Excel=Microsoft.Office.Interop.Excel;
using log4net;
using System.Diagnostics;
using System.Linq;
using System;
using System.ComponentModel;
using System.Timers;
using System.Collections.Concurrent;
using System.Threading;

namespace xltrail.Client
{
    public static class Addin
    {
        [ComVisible(true)]
        public class RibbonController : ExcelRibbon, IExcelAddIn
        {
            static Excel.Application xlApp;
            static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            static BackgroundWorker worker;
            static ConcurrentQueue<string> repositories;
            IRibbonUI ribbon;


            public void AutoOpen()
            {
                var loggerPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), "xltrail", "logs");
                if (!Directory.Exists(loggerPath))
                    Directory.CreateDirectory(loggerPath);
                Logger.Setup();
                logger.Info("Initialise Addin");
                xlApp = (Excel.Application)ExcelDnaUtil.Application;
                xlApp.WorkbookOpen += XlApp_WorkbookOpen;
                xlApp.WorkbookAfterSave += XlApp_WorkbookAfterSave;
                xlApp.WorkbookActivate += XlApp_WorkbookActivate;

                repositories = new ConcurrentQueue<string>();
                var worker = new Thread(Consume);
                worker.Start();
            }

            private void XlApp_WorkbookActivate(Excel.Workbook workbook)
            {
                var path = workbook.FullName;
                var directory = Path.GetDirectoryName(path);
                if (!LibGit2Sharp.Repository.IsValid(directory))
                {
                    xlApp.Caption = null;
                }
                else
                {
                    Refresh();
                }
            }

            private static void Refresh()
            {
                var path = xlApp.ActiveWorkbook.FullName;
                var directory = Path.GetDirectoryName(path);
                var repository = new LibGit2Sharp.Repository(directory);
                xlApp.Caption = repository.Head.FriendlyName + " [" + repository.Head.Tip.Id.Sha.Substring(0, 7) + "]";

            }

            private static void Push(string directory)
            {
                //libgit2 repo
                var repository = new LibGit2Sharp.Repository(directory);
                var branch = repository.Head;

                //git command line wrapper
                var git = new Git(directory);

                //push to origin (if exists)
                var origin = repository.Network.Remotes.Where(x => x.Name == "origin").FirstOrDefault();
                var trackedBranch = repository.Branches.Where(b => b.FriendlyName == "origin/" + repository.Head.FriendlyName).FirstOrDefault();
                if (origin != null && trackedBranch != null)
                {
                    //fetch from origin
                    git.Fetch();

                    //calculate divergence between origin and local branch
                    var divergence = repository.ObjectDatabase.CalculateHistoryDivergence(repository.Head.Tip, trackedBranch.Tip);

                    //rebase if local is behind origin
                    if (divergence.BehindBy > 0)
                    {
                        logger.InfoFormat("Reset branch {0}", branch.FriendlyName);
                        repository.Reset(LibGit2Sharp.ResetMode.Hard, branch.TrackedBranch.Tip);
                    }

                    //push if local is ahead of origin
                    if (divergence.AheadBy > 0)
                    {
                        var url = origin.PushUrl;
                        logger.InfoFormat("{0} is behind {1} by {2} commit(s)", directory, url, branch.TrackingDetails.AheadBy);
                        logger.InfoFormat("Push {0} to {1}", directory, url);
                        git.Push(branch.FriendlyName);
                    }
                }
            }

            private static void Consume()
            {
                logger.Info("Start consumer thread");
                string directory;
                while (true)
                {
                    if (repositories.TryDequeue(out directory))
                    {
                        logger.InfoFormat("New item in queue: {0}", directory);

                        try
                        {
                            Push(directory);
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message, e);
                            ShowNotification(e.Message);
                        }
                    }
                    Thread.Sleep(2000);
                }
            }


            private void XlApp_WorkbookOpen(Excel.Workbook workbook)
            {
                xlApp.ScreenUpdating = true;
                xlApp.StatusBar = null;
                var path = workbook.FullName;
                logger.InfoFormat("Open workbook {0}", path);

                var directory = Path.GetDirectoryName(path);
                if (!LibGit2Sharp.Repository.IsValid(directory))
                    return;

                logger.InfoFormat("{0} is a valid Git repository", directory);

                //workbook path relative to repository
                var gitPath = path.Substring(directory.Length + 1).Replace("\\", "/");

                //libgit2 repo
                var repository = new LibGit2Sharp.Repository(directory);

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
                    logger.InfoFormat("Fetch from remote/origin: {0}", url);
                    xlApp.StatusBar = "Fetching from " + url;
                    git.Fetch();
                    xlApp.StatusBar = null;

                    var trackedBranch = repository.Branches.Where(b => b.IsRemote).Where(b => b.FriendlyName == "origin/" + repository.Head.FriendlyName).FirstOrDefault();

                    //calculate divergence between origin and local branch
                    var divergence = repository.ObjectDatabase.CalculateHistoryDivergence(repository.Head.Tip, trackedBranch.Tip);

                    //check if local is behind origin
                    if (divergence.BehindBy > 0)
                    {
                        logger.InfoFormat("Newer commit available on from remote/origin: {0}", trackedBranch.Tip.Id);
                        xlApp.StatusBar = "Pulling newer " + gitPath + " version from " + url;
                        var localWorkbook = repository.Head[gitPath];
                        var remoteWorkbook = trackedBranch[gitPath];

                        var localFileSha = ((LibGit2Sharp.Blob)localWorkbook.Target).Sha;
                        var remoteFileSha = ((LibGit2Sharp.Blob)remoteWorkbook.Target).Sha;

                        xlApp.ScreenUpdating = false;

                        //close active workbook
                        logger.InfoFormat("Close workbook {0} to update branch from remote/origin", gitPath);
                        workbook.Close(false);

                        //reset local to remote
                        //TODO: this should only apply to current file (and not the entire branch)
                        logger.InfoFormat("Reset branch {0}", repository.Head.FriendlyName);
                        repository.Reset(LibGit2Sharp.ResetMode.Hard, trackedBranch.Tip);

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

                //git command line wrapper
                var git = new Git(directory);

                logger.InfoFormat("Commit new workbook version: {0}", path);

                //add
                LibGit2Sharp.Commands.Stage(repository, path);

                //commit
                git.Commit("Modified " + name);

                //add to list for background push
                if (!repositories.Contains(directory))
                    repositories.Enqueue(directory);

                Refresh();
            }

            private static void ShowNotification(string description)
            {
                var notification = new System.Windows.Forms.NotifyIcon()
                {
                    Visible = true,
                    Icon = System.Drawing.SystemIcons.Information,
                    BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning,
                    BalloonTipTitle = "xltrail",
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
