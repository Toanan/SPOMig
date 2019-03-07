using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.SharePoint.Client;
using System.IO;
using System.ComponentModel;
using SPOMig.Windows;

namespace SPOMig
{
    /// <summary>
    /// Window to select a local path to copy and a SharePoint Online library as a target
    /// </summary>
    public partial class Migration : Window
    {
        //We instanciate the backgroundWorker to copy folders and files 
        BackgroundWorker bw = new BackgroundWorker();

        #region Props
        public string LocalPath { get; set; }
        public ClientContext Context { get; set; }
        public string DocLib { get; set; }
        public bool DeleteOldItems { get; set; }
        public int BatchRequestSize { get; set; }
        #endregion

        #region Ctor
        public Migration(ListCollection ListCollection, ClientContext ctx)
        {
            InitializeComponent();
            this.Context = ctx;

            this.BatchRequestSize = Convert.ToInt32(FileLogic.getXMLSettings("BatchRequestSize"));

            //We retrieve only document library from the ListCollection passed and populate the combobox
            foreach (List list in ListCollection)
            {
                if (list.BaseTemplate == 101)
                    Cb_doclib.Items.Add(list.Title);
            }

            //We set the top header of the window
            Lb_Top.Content = "Select a local path and a library to synchronize";
            Img_SP.Visibility = Visibility.Visible;
            Cb_doclib.SelectedIndex = 0;
            Lb_Connected.Content = $"Connected to {ctx.Url}";

            //We set the backgroundWorker Delegates
            bw.DoWork += bw_Dowork;
            bw.WorkerReportsProgress = true;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
            bw.ProgressChanged += bw_ProgressChanged;

            bw.WorkerSupportsCancellation = true;
        }

        public Migration(List odList, ClientContext ctx)
        {
            InitializeComponent();
            this.Context = ctx;

            this.BatchRequestSize = Convert.ToInt32(FileLogic.getXMLSettings("BatchRequestSize"));

            Cb_doclib.Items.Add(odList.Title);
            Cb_doclib.SelectedIndex = 0;

            //We set the top header of the window
            Lb_Top.Content = "Select a local path to synchronize";
            Img_OD.Visibility = Visibility.Visible;
            Lb_Connected.Content = $"Connected to {ctx.Url}";

            //We set the backgroundWorker Delegates
            bw.DoWork += bw_Dowork;
            bw.WorkerReportsProgress = true;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
            bw.ProgressChanged += bw_ProgressChanged;

            bw.WorkerSupportsCancellation = true;
        }
        #endregion

        #region EventHandler

        #region BgWorker

        /// <summary>
        /// BackgroundWorker Work event:
        /// Copy file and folders from the path selected to the target SharePoint Online library
        /// </summary>
        /// <param name="sender">Btn_Copy</param>
        /// <param name="e">OnClick()</param>
        private void bw_Dowork(object sender, DoWorkEventArgs e)
        {
            //[RESULT/LOG] We instanciate the repporting object
            Reporting repport = new Reporting(this.DocLib, this.Context.Web.ServerRelativeUrl);
            //[LOG:Verbose] We create the log object and log Local Path formating
            CopyLog log = new CopyLog(CopyLog.Status.Verbose, "Local path formating", LocalPath, "");
            repport.writeLog(log);

            try
            {
                //We ensure the localpath endwith "/" for further formating actions
                if (!this.LocalPath.EndsWith("/") || !this.LocalPath.EndsWith("\\")) this.LocalPath = this.LocalPath + "\\";

                //[LOG:OK] Local Path formating : Log success
                log.ActionStatus = CopyLog.Status.OK;
                repport.writeLog(log);

                //[LOG:Verbose] Local file retrieve
                log.update(CopyLog.Status.Verbose, "Local file retrieve", LocalPath, "");
                repport.writeLog(log);

                //We retrive the list of DirectoryInfo and FileInfo
                List<FileInfo> files = FileLogic.getFiles(LocalPath);
                List<DirectoryInfo> folders = FileLogic.getSourceFolders(LocalPath);

                //[LOG:OK] Local file retrieve
                log.ActionStatus = CopyLog.Status.OK;
                repport.writeLog(log);

                //We instanciate the SPOLogic object to interact with SharePoint Online
                SPOLogic ctx = new SPOLogic(Context);

                //[LOG:Verbose] Checking library
                log.update(CopyLog.Status.Verbose, "Checking library", LocalPath, "");
                repport.writeLog(log);

                //We enable Folder creation for the SharePoint Online library and ensure the Hash column exist
                bw.ReportProgress(0, "Checking Library");
                List list = ctx.setLibraryReadyForPRocessing(this.DocLib);

                //[LOG:OK] Checking library
                log.ActionStatus = CopyLog.Status.OK;
                repport.writeLog(log);

                //[LOG:Verbose] Online ListItem retrieve
                log.update(CopyLog.Status.Verbose, "Online ListItem retrieve", LocalPath, "");
                repport.writeLog(log);

                // We retrieve all listitems in the library
                bw.ReportProgress(0, "Retrieving ListItems");
                List<ListItem> onlineListItem = ctx.GetAllDocumentsInaLibrary(this.DocLib);

                //[LOG:OK] Online ListItem retrieve
                log.ActionStatus = CopyLog.Status.OK;
                repport.writeLog(log);

                #region Folder Creation

                //[LOG:Title] Folder creation beggins
                log.update("[Starting Folder Creation process]");
                repport.writeLog(log);

                var rootFolder = list.RootFolder;
                Context.Load(rootFolder);
                Context.ExecuteQuery();

                //We set the index to display progression
                int i = 0;
                int count = 0;

                foreach (DirectoryInfo folder in folders)
                {
                    //Progression display
                    i++;
                    double percentage = (double)i / folders.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Checking folders - {advancement}%\n{i}/{folders.Count}");

                    //We check for pending cancellation
                    if (bw.CancellationPending == true)
                    {
                        //[LOG:CANCEL] Cancellation log
                        log.update("[Process Cancelled]");
                        repport.writeLog(log);

                        //We cancel the backgroundWorker and return
                        e.Cancel = true;
                        return;
                    }
                    //If no cancellation, we launch the copy folder process
                    else
                    {
                        //[LOG:Verbose] Folder Creation
                        log.update(CopyLog.Status.Verbose, "Folder creation", folder.FullName, "");
                        repport.writeLog(log);

                        //We process the folder
                        CopyStatus copyStatus = ctx.copyFolderToSPO(folder, list, LocalPath, onlineListItem);
                                              

                        //[LOG:OK] Folder Creation
                        log.ActionStatus = CopyLog.Status.OK;
                        if (copyStatus != null)
                        {
                            //[LOG:OK] Folder Creation update path
                            log.ItemPath = copyStatus.Path;
                            log.Comment = copyStatus.Comment;
                            repport.writeLog(log);
                            //[RESULT] Folder Creation
                            repport.writeResult(copyStatus);
                        }
                        else
                        {
                            //We skip writing result for the rootfolder
                            repport.writeLog(log);
                        }
                    }
                }


                i = 0;
                count = 0;
                foreach (FoldersToProcess folder in ctx.FoldersToUpload)
                {

                    //Progression display
                    i++;
                    double percentage = (double)i / folders.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Batching folders upload - {advancement}%\n{i}/{folders.Count}");

                    count++;
                    var myFolder = rootFolder.Folders.Add(folder.ItemUrls.ServerRelativeUrl);
                    if (count >= this.BatchRequestSize)
                    {
                        bw.ReportProgress(advancement, "Uploading folders batch");
                        Context.RequestTimeout = -1;
                        Context.ExecuteQuery();
                        count = 0;
                    }
                }
                bw.ReportProgress(0, "Finalising folders upload");
                Context.RequestTimeout = -1;
                Context.ExecuteQuery();

                i = 0;
                count = 0;

                foreach (FoldersToProcess folder in ctx.FoldersToUpload)
                {

                    //Progression display
                    i++;
                    double percentage = (double)i / folders.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Batching folders metadata - {advancement}%\n{i}/{folders.Count}");

                    //We update metadate
                    ListItem listitemFolder = Context.Web.GetListItem(folder.ItemUrls.ServerRelativeUrl);
                    listitemFolder["Created"] = folder.Created;
                    listitemFolder["Modified"] = folder.Modified;
                    listitemFolder.Update();
                    count++;
                    if (count >= this.BatchRequestSize)
                    {
                        bw.ReportProgress(advancement, "Uploading folders metadata batch");
                        Context.RequestTimeout = -1;
                        Context.ExecuteQuery();
                        count = 0;
                    }
                }
                
                bw.ReportProgress(0, "Finalising folders Metadata");
                Context.RequestTimeout = -1;
                Context.ExecuteQuery();
                #endregion


                #region FileCopy

                //[LOG:Title] File upload beggins
                log.update("[Starting File Upload process]");
                repport.writeLog(log);

                //We reset the progression index
                i = 0;

                foreach (FileInfo file in files)
                {
                    //Progression display
                    i++;
                    double percentage = (double)i / files.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Checking files - {advancement}%\n{i}/{files.Count}");

                    //Check if Cancellation is pending
                    if (bw.CancellationPending == true)
                    {
                        //[LOG:CANCEL] Cancellation log
                        log.update("[Process Cancelled]");
                        repport.writeLog(log);

                        //We cancel the backgroundWorker and return
                        e.Cancel = true;
                        return;
                    }
                    //If no cancellation, we launch the copy file process
                    else
                    {
                        //[LOG:Verbose] File Upload
                        log.update(CopyLog.Status.Verbose, "File upload", file.FullName, "");
                        repport.writeLog(log);

                        //We copy the file
                        CopyStatus copyStatus = ctx.copyFileToSPO(file, list, LocalPath, onlineListItem);

                        //[LOG:OK] File Upload
                        log.ActionStatus = CopyLog.Status.OK;
                        log.ItemPath = copyStatus.Path;
                        log.Comment = copyStatus.Comment;

                        //[RESULT/LOG: OK] File Upload
                        repport.writeLog(log);
                        repport.writeResult(copyStatus);
                    }
                }
                #endregion

                #region Item Deletion

                //We handle the delete old items checkbox feature
                if (this.DeleteOldItems == true)
                {
                    //[LOG:Verbose] Online ListItem retrieve
                    log.update(CopyLog.Status.Verbose, "Online ListItem retrieve", LocalPath, "");
                    repport.writeLog(log);

                    // We retrieve all listitems in the library and divide files and folders
                    bw.ReportProgress(0, "Retrieving ListItems");
                    onlineListItem = ctx.GetAllDocumentsInaLibrary(this.DocLib);
                    List<ListItem> onlineFiles = ctx.GetOnlyFiles(onlineListItem);
                    List<ListItem> onlineFolders = ctx.GetOnlyFolders(onlineListItem);

                    //[LOG:OK] Online ListItem retrieve
                    log.ActionStatus = CopyLog.Status.OK;
                    repport.writeLog(log);

                    //[LOG:Title] File deletion beggins
                    log.update("[Starting File Deletion process]");
                    repport.writeLog(log);

                    #region File Deletion

                    //We retrieve all the formated urls from local source files
                    List<ItemURLs> itemsUrls = new List<ItemURLs>();
                    foreach (FileInfo file in files)
                    {
                        ItemURLs itemUrl = ctx.formatUrl(file, list, LocalPath);
                        itemsUrls.Add(itemUrl);
                    }

                    //We reset the progression index
                    i = 0;
                    count = 0;
                    foreach (ListItem onlineFile in onlineFiles)
                    {

                        //Progression display
                        i++;
                        double percentage = (double)i / onlineFiles.Count;
                        int advancement = Convert.ToInt32(percentage * 100);
                        bw.ReportProgress(advancement, $"Checking old files - {advancement}%\n{i}/{onlineFiles.Count}");

                        //Check if Cancellation is pending
                        if (bw.CancellationPending == true)
                        {
                            //[LOG:CANCEL] Cancellation log
                            log.update("[Process Cancelled]");
                            repport.writeLog(log);

                            //We cancel the backgroundWorker and return
                            e.Cancel = true;
                            return;
                        }
                        //If no cancellation, we launch the file deletion process
                        else
                        {

                            //[LOG:Verbose] File Deletion
                            log.update(CopyLog.Status.Verbose, "File deletion", (string)onlineFile["FileLeafRef"], "");
                            repport.writeLog(log);

                            //Attempt to delete the file if necessary
                            CopyStatus copystat = ctx.CheckItemToDelete(itemsUrls, list, LocalPath, onlineFile);

                            
                            if (copystat.Status == CopyStatus.ItemStatus.Deleted)
                            {
                                count ++;
                                ListItem item = list.GetItemById((Int32)onlineFile["ID"]);
                                item.DeleteObject();
                            }
                            if (count >= this.BatchRequestSize)
                            {
                                count = 0;
                                Context.ExecuteQuery();
                            }

                            //[LOG:OK] File Deletion
                            log.ActionStatus = CopyLog.Status.OK;
                            log.ItemPath = copystat.Path;
                            log.Comment = copystat.Comment;

                            //[RESULT/LOG: OK] File Deletion
                            repport.writeLog(log);
                            repport.writeResult(copystat);

                        }
                        
                    }
                    bw.ReportProgress(0, "Deleting old files");
                    Context.RequestTimeout = -1;
                    Context.ExecuteQuery();

                    #endregion

                    #region Folder Deletion

                    //Folder deletion

                    //We retrieve all the formated urls from local source folders
                    List<ItemURLs> folderUrls = new List<ItemURLs>();
                    foreach (DirectoryInfo folder in folders)
                    {
                        ItemURLs itemUrl = ctx.formatUrl(folder, list, LocalPath);
                        folderUrls.Add(itemUrl);
                    }

                    //We reset the progression index
                    i = 0;
                    count = 0;
                    foreach (ListItem onlineFolder in onlineFolders)
                    {

                        //Progression display
                        i++;
                        Double percentage = (double)i / onlineFolders.Count;
                        int advancement = Convert.ToInt32(percentage * 100);
                        bw.ReportProgress(advancement, $"Checking old folders - {advancement}%\n{i}/{onlineFolders.Count}");

                        //Check if Cancellation is pending
                        if (bw.CancellationPending == true)
                        {
                            //[LOG:CANCEL] Cancellation log
                            log.update("[Process Cancelled]");
                            repport.writeLog(log);

                            //We cancel the backgroundWorker and return
                            e.Cancel = true;
                            return;
                        }
                        //If no cancellation, we launch the folder deletion process
                        else
                        {

                            //[LOG:Verbose] Folder Deletion
                            log.update(CopyLog.Status.Verbose, "Folder deletion", (string)onlineFolder["FileLeafRef"], "");
                            repport.writeLog(log);

                            //Attempt to delete the file if necessary
                            CopyStatus copystat = ctx.CheckItemToDelete(folderUrls, list, LocalPath, onlineFolder);
                            //Change the item type to folder for result purpose
                            copystat.Type = CopyStatus.ItemType.Folder;

                            int pathRootFolderCount = FileLogic.isRootFolder(list.RootFolder.ServerRelativeUrl);
                            int pathFolderCount = FileLogic.isRootFolder(copystat.Path);

                            if ((pathFolderCount - pathRootFolderCount) == 1 && copystat.Status == CopyStatus.ItemStatus.Deleted)
                            {
                                count++;
                                ListItem item = list.GetItemById((Int32)onlineFolder["ID"]);
                                item.DeleteObject();
                            }
                            if (count >= this.BatchRequestSize)
                            {
                                Context.RequestTimeout = -1;
                                try
                                {
                                    Context.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {
                                    if (!ex.Message.Contains("Item does not exist. It may have been deleted by another user."))
                                    {
                                        throw ex;
                                    }
                                }
                            }

                            //[LOG:OK] Folder Deletion
                            log.ActionStatus = CopyLog.Status.OK;
                            log.ItemPath = copystat.Path;
                            log.Comment = copystat.Comment;

                            //[RESULT/LOG: OK] Folder Deletion
                            repport.writeLog(log);
                            repport.writeResult(copystat);
                           
                        }
                    }
                    bw.ReportProgress(0, "Deleting old folders");
                    Context.RequestTimeout = -1;
                    try
                    {
                        Context.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        if (!ex.Message.Contains("Item does not exist. It may have been deleted by another user."))
                        {
                            throw ex;
                        }
                    }
                    #endregion
                }

                #endregion

            }
            catch (Exception ex)
            {
                //We cancel the backgroundWorker if an exception is not handled so far
                MessageBox.Show(ex.Message);
                bw.CancelAsync();
                e.Cancel = true;

                //[LOG:ERROR] Error log
                log.ActionStatus = CopyLog.Status.Error;
                log.Comment = ex.Message;
                repport.writeLog(log);

                //[RESULT:ERROR] Error result
                CopyStatus copyError = new CopyStatus
                {
                    Comment = ex.Message,
                    Name = log.ItemPath,
                    Path = log.ItemPath,
                    Status = CopyStatus.ItemStatus.Error
                };
                repport.writeResult(copyError);
            }
        }

        /// <summary>
        /// BackgroundWorker Completion Event :
        /// Show Migration Window if work was successfull
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Pb_progress.Value = e.ProgressPercentage;
            Lb_State.Content = e.UserState;
        }

        /// <summary>
        /// BackgroundWorker Completion Event :
        /// Show Migration Window if work was successfull
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //UI update
            this.IsEnabled = true;
            Btn_Cancel.IsEnabled = true;
            Pb_progress.Visibility = Visibility.Hidden;
            Btn_Cancel.Visibility = Visibility.Hidden;
            Btn_Copy.IsEnabled = true;
            Tb_LocalPath.IsEnabled = true;
            Btn_Clean.IsEnabled = true;
            Btn_Home.IsEnabled = true;
            Cb_doclib.IsEnabled = true;
            Chbx_deleteItems.IsEnabled = true;
            Pb_progress.Value = 0;
            Lb_State.Content = "";

            //If no cancelation, we show a message box and open the relust file path
            if (e.Cancelled == false)
            {
                MessageBox.Show("Task finished successfully");
                string appPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string appName = "SPOMig";
                System.Diagnostics.Process.Start($"{appPath}/{appName}/Results/");
                bw.Dispose();
            }
            bw.Dispose();
        }

        #endregion

        #region Button
        
        /// <summary>
        /// Button Copy OnClick event :
        /// Input check then BackgroundWorker launch
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (Tb_LocalPath.Text.Contains(".csv"))
            {
                this.DocLib = Cb_doclib.Text;
                this.DeleteOldItems = (bool)Chbx_deleteItems.IsChecked;

                if (!Directory.Exists(Tb_LocalPath.Text))
                {
                    MessageBox.Show("Cannot find local path, please double check");
                    return;
                }



            }
            else
            {
                this.DocLib = Cb_doclib.Text;
                this.DeleteOldItems = (bool)Chbx_deleteItems.IsChecked;


                this.LocalPath = @Tb_LocalPath.Text;

                #region UserInput checks
                if (LocalPath == "")
                {
                    MessageBox.Show("Please fill the local path field");
                    return;
                }
                if (!Directory.Exists(LocalPath))
                {
                    MessageBox.Show("Cannot find local path, please double check");
                    return;
                }
                /*
                if (Cb_doclib.SelectedItem == null)
                {
                    MessageBox.Show("Please select a document library");
                    return;
                }*/
                #endregion

                //UI update
                Btn_Copy.IsEnabled = false;
                Tb_LocalPath.IsEnabled = false;
                Btn_Clean.IsEnabled = false;
                Btn_Home.IsEnabled = false;
                Cb_doclib.IsEnabled = false;
                Chbx_deleteItems.IsEnabled = false;
                Pb_progress.Value = 0;
                Btn_Cancel.Visibility = Visibility.Visible;
                Pb_progress.Visibility = Visibility.Visible;

                //Run the BackgroundWorker
                bw.RunWorkerAsync();
            }  
        }

        /// <summary>
        /// Button Cancel OnClick event : Cancel the BackgroundWorker
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            bw.CancelAsync();
            Btn_Cancel.IsEnabled = false;
            this.IsEnabled = false;
        }

        #endregion

        #endregion

        /// <summary>
        /// Buttom Home to navigate to the login window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnHome_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            Startup mw = new Startup();
            mw.Show();
            bw.Dispose();
            this.Close();
        }

        /// <summary>
        /// Button Clean Library to clear library from the Hash column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClean_Click(object sender, RoutedEventArgs e)
        {
            //We update the UI
            this.IsEnabled = false;
            string libName = Cb_doclib.Text;

            //We create the context object and call the column supression method
            SPOLogic spo = new SPOLogic(Context);
            bool didClean = spo.cleanLibraryFromProcessing(libName);

            if (didClean) MessageBox.Show("Library is clean");

            //We reactivate the UI
            this.IsEnabled = true;

        }

    }
}
