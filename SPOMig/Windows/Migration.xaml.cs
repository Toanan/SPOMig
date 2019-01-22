﻿using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.SharePoint.Client;
using System.IO;
using System.ComponentModel;

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
        #endregion

        #region Ctor
        public Migration(ListCollection ListCollection, ClientContext ctx)
        {
            InitializeComponent();
            this.Context = ctx;
            

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
            Reporting repport = new Reporting(this.DocLib);
            //[LOG:Verbose] We create the log object and log Local Path formating
            CopyLog log = new CopyLog(CopyLog.Status.Verbose, "Local path formating", LocalPath, "");
            repport.writeLog(log);

            try
            {
                //We ensure the localpath endwith "/" for further formating actions
                if (!this.LocalPath.EndsWith("/") ||!this.LocalPath.EndsWith("\\")) this.LocalPath = this.LocalPath + "\\";

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
                List list = ctx.setLibraryReadyForPRocessing(this.DocLib);

                //[LOG:OK] Checking library
                log.ActionStatus = CopyLog.Status.OK;
                repport.writeLog(log);

                #region Folder Creation

                //[LOG:Title] Folder creation beggins
                log.update("[Starting Folder Creation process]");
                repport.writeLog(log);

                //We set the index to display progression
                int i = 0;

                foreach (DirectoryInfo folder in folders)
                {
                    //Progression display
                    i++;
                    double percentage = (double)i / folders.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Copying folders {advancement}%\n{i}/{folders.Count}");

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
                        CopyStatus copyStatus = ctx.copyFolderToSPO(folder, list, LocalPath);

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
                    bw.ReportProgress(advancement, $"Copying files {advancement}%\n{i}/{files.Count}");

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
                        CopyStatus copyStatus = ctx.copyFileToSPO(file, list, LocalPath);

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
            Cb_doclib.IsEnabled = true;
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
            this.LocalPath = @Tb_LocalPath.Text;
            this.DocLib = Cb_doclib.Text;

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
            Cb_doclib.IsEnabled = false;
            Pb_progress.Value = 0;
            Btn_Cancel.Visibility = Visibility.Visible;
            Pb_progress.Visibility = Visibility.Visible;

            //Run the BackgroundWorker
            bw.RunWorkerAsync();  
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
            MainWindow mw = new MainWindow();
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
            bool didClean = spo.clanLibraryFromProcessing(libName);

            if (didClean) MessageBox.Show("Library is clean");

            //We reactivate the UI
            this.IsEnabled = true;

        }
    }
}
