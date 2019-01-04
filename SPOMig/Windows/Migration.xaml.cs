using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using System.IO;
using System.ComponentModel;

namespace SPOMig
{
    /// <summary>
    /// Logique d'interaction pour Migration.xaml
    /// </summary>
    public partial class Migration : Window
    {

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
            // Retrive only document library from the ListCollection passed
            foreach (List list in ListCollection)
            {
                if (list.BaseTemplate == 101)
                    Cb_doclib.Items.Add(list.Title);
            }

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
            Pb_progress.Visibility = Visibility.Hidden;
            Btn_Cancel.Visibility = Visibility.Hidden;
            Btn_Copy.IsEnabled = true;
            Pb_progress.Value = 0;
            Lb_State.Content = "";

            if (e.Cancelled == false)
            {
                MessageBox.Show("Task finished successfully");  
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
        /// Retrive user info to retrive the SPO Site lists
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_Dowork(object sender, DoWorkEventArgs e)
        {
           
            try
            {
                //We append "/" to LocalPath for formating purpose
                if (!this.LocalPath.EndsWith("/")) this.LocalPath = this.LocalPath + "/";

                //Retrieving Local files and folders
                FileLogic FileLogic = new FileLogic(LocalPath);
                List<FileInfo> files = FileLogic.getFiles();
                List<DirectoryInfo> folders = FileLogic.getSourceFolders();

                //We load the library
                List list = Context.Web.Lists.GetByTitle(this.DocLib);
                
                //Enable Folder creation for the library
                list.EnableFolderCreation = true;
                list.Update();
                Context.Load(list.RootFolder);
                Context.ExecuteQuery();

                SPOLogic ctx = new SPOLogic(Context);
                
                //Folders Copy
                int i = 0;
                foreach (DirectoryInfo folder in folders)
                {
                    i++;
                    double percentage = (double)i/folders.Count;
                    int advancement = Convert.ToInt32(percentage*100);
                    bw.ReportProgress(advancement, $"Copying folders {advancement}%\n{i}/{folders.Count}");
                    
                    //Handle cancellation
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        try
                        {
                            ctx.copyFolderToSPO(folder, list, LocalPath);
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.EndsWith("already exists."))
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                    }  
                }

                //Files copy
                i = 0;
                foreach (FileInfo file in files)
                {
                    i++;
                    double percentage = (double)i / files.Count;
                    int advancement = Convert.ToInt32(percentage * 100);
                    bw.ReportProgress(advancement, $"Copying files {advancement}%\n{i}/{files.Count}");

                    //Handle cancellation
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        ctx.copyFileToSPO(file, list, LocalPath);
                    }    
                }
            }
            catch (Exception ex)
            {
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

            Btn_Copy.IsEnabled = false;
            Pb_progress.Value = 0;
            Btn_Cancel.Visibility = Visibility.Visible;
            Pb_progress.Visibility = Visibility.Visible;

            
            //BackgroundWorker Delegates
            bw.DoWork += bw_Dowork;
            bw.WorkerReportsProgress = true;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
            bw.ProgressChanged += bw_ProgressChanged;

            bw.WorkerSupportsCancellation = true;
            bw.RunWorkerAsync();
            
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            bw.CancelAsync();
            Btn_Cancel.IsEnabled = false;
        }

        #endregion

        #endregion

    }
}
