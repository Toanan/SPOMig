using System;
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
        public string RepportPath { get; set; }
        #endregion

        #region Ctor
        public Migration(ListCollection ListCollection, ClientContext ctx)
        {
            InitializeComponent();
            this.Context = ctx;
            // Retrive only document library from the ListCollection passed
            foreach (List list in ListCollection)
            {
                if (list.BaseTemplate == 101)
                    Cb_doclib.Items.Add(list.Title);
            }
            //BackgroundWorker Delegates
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
        /// BackgroundWorker Completion Event :
        /// Show Migration Window if work was successfull
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.IsEnabled = true;
            Btn_Cancel.IsEnabled = true;
            Pb_progress.Visibility = Visibility.Hidden;
            Btn_Cancel.Visibility = Visibility.Hidden;
            Btn_Copy.IsEnabled = true;
            Pb_progress.Value = 0;
            Lb_State.Content = "";

            if (e.Cancelled == false)
            {
                MessageBox.Show("Task finished successfully");
                System.Diagnostics.Process.Start($"{AppDomain.CurrentDomain.BaseDirectory}/Results/");
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
        /// BackgroundWorker Work event:
        /// Copy file and folders to the target SharePoint Online library
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_Dowork(object sender, DoWorkEventArgs e)
        {
            try
            {
                WriteRepport repport = new WriteRepport(this.DocLib);
                this.RepportPath = repport.ResultFilePath; 

                //We append "/" to LocalPath for formating purpose
                if (!this.LocalPath.EndsWith("/")) this.LocalPath = this.LocalPath + "/";

                //We retrive Local files and folders using the FileLogic class
                FileLogic FileLogic = new FileLogic(LocalPath);
                List<FileInfo> files = FileLogic.getFiles();
                List<DirectoryInfo> folders = FileLogic.getSourceFolders();

                //We instanciate the SPOLogic class to copy folders and file
                SPOLogic ctx = new SPOLogic(Context);

                //We enable Folder creation for the SharePoint Online library and ensure the Hash column exist
                List list = ctx.setLibraryReadyForPRocessing(this.DocLib);

                //We set the index to display progression
                int i = 0;

                #region FolderCopy
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
                        e.Cancel = true;
                    }
                    else
                    {
                        //If no cancellation, we try to copy the folder
                        bool didCopy = ctx.copyFolderToSPO(folder, list, LocalPath);
                        if (didCopy)
                        {
                            repport.writeResult($"{folder.Name},Folder,{folder.FullName.Remove(0, LocalPath.Length)},Copied,");
                            //"Name,Type,OnlinePath,Status,Comment"
                        }
                        else
                        {
                            repport.writeResult($"{folder.Name},Folder,{folder.FullName.Remove(0, LocalPath.Length)},Allready exists,");
                        }
                    }
                }
                #endregion


                #region FileCopy
                i = 0;
                foreach (FileInfo file in files)
                {
                    i++;
                    double percentage = (double)i / files.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Copying files {advancement}%\n{i}/{files.Count}");

                    //Check if Cancellation is pending
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        bool didCopy = ctx.copyFileToSPO(file, list, LocalPath);

                        if (didCopy)
                        {
                            repport.writeResult($"{file.Name},File,{file.FullName.Remove(0,LocalPath.Length)},Copied,");
                            //"Name,Type,OnlinePath,Status,Comment"
                        }
                        else
                        {
                            repport.writeResult($"{file.Name},File,{file.FullName.Remove(0, LocalPath.Length)},Allready exist,");
                            //"Name,Type,OnlinePath,Status,Comment"
                        }
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
            }
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
            if (Cb_doclib.SelectedItem == null)
            {
                MessageBox.Show("Please select a document library");
                return;
            }
            #endregion

            //UI update
            Btn_Copy.IsEnabled = false;
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

    }
}
