using System;
using System.Security;
using System.Windows;
using System.ComponentModel;
using Microsoft.SharePoint.Client;

namespace SPOMig
{
    /// <summary>
    /// Class designed to handle the connection to a SharePoint Online Site
    /// </summary>
    public partial class MainWindow : Window
    {

        BackgroundWorker bw = new BackgroundWorker();

        #region Props
        public string SiteUrl { get; set; }
        public string UserName { get; set; }
        public SecureString PassWord { get; set; }
        public ListCollection Lists { get; set; }
        public ClientContext Context { get; set; }
        public List ODLibrary { get; set; }
        #endregion

        #region Ctor
        public MainWindow()
        {
            InitializeComponent();
            Lb_Top.Content = "First connect to a OneDrive or SharePoint Online Site";  
        }
        #endregion

        #region EventHandlers

        #region BgWorker

        /// <summary>
        /// BackgroundWorker Completion Event :
        /// Show Migration Window if work was successfull
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //If the backgroundWorker complete the task successfully, we open themigration window
            if (e.Cancelled == false)
            {
                this.Hide();
                if (ODLibrary != null)
                {
                    Migration mig = new Migration(ODLibrary, Context);
                    mig.Show();
                }
                else
                {
                    Migration mig = new Migration(Lists, Context);
                    mig.Show();
                }
                bw.Dispose();
                this.Close();
            }
            else
            {
                //We reactivate the UI
                Btn_Connect.IsEnabled = true;
                Tb_SPOSite.IsEnabled = true;
                Tb_UserName.IsEnabled = true;
                Pb_PassWord.IsEnabled = true;
            }
        }

        /// <summary>
        /// BackgroundWorker Work event:
        /// Retrive user info to retrive the SPO Site lists
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bw_Dowork(object sender, DoWorkEventArgs e)
        {
            //We try to connect to the SharePoint Online site and retrieve the libraries
            try
            {
                using (var ctx = new ClientContext(SiteUrl))
                {
                    SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(UserName, PassWord);
                    ctx.Credentials = credentials;

                    ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                    this.Context = ctx;

                    SPOLogic spol = new SPOLogic(ctx);

                    //We check if the SPO site is a OneDrive Url, and process accordingly
                    if (SiteUrl.Contains("/personal/"))
                    {
                        this.ODLibrary = spol.getODList();
                    }
                    //Else we have a SPO Site Url, and process accordingly
                    else
                    {
                        this.Lists = spol.getLists();
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
        /// Button Connect OnClick event :
        /// Input check then BackgroundWorker launch
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            #region UserInput Checks
            if (Tb_SPOSite.Text == "")
            {
                MessageBox.Show("Please fill the site URL field");
                return;
            }
            if (Tb_UserName.Text == "")
            {
                MessageBox.Show("Please fill the User Name field");
                return;
            }
            if (Pb_PassWord.Password == "")
            {
                MessageBox.Show("Please fill the Pass Word field");
                return;
            }
            #endregion

            //BarckgroundWorker delegates
            bw.DoWork += bw_Dowork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;

            //UI update
            Btn_Connect.IsEnabled = false;
            Tb_SPOSite.IsEnabled = false;
            Tb_UserName.IsEnabled = false;
            Pb_PassWord.IsEnabled = false;

            //We set the class properties
            this.SiteUrl = Tb_SPOSite.Text;
            this.UserName = Tb_UserName.Text;
            this.PassWord = Pb_PassWord.SecurePassword;
            
            //We ensure the backgroundWorker supports cancellation and run it
            bw.WorkerSupportsCancellation = true;
            bw.RunWorkerAsync();
        }

        #endregion

        #endregion

    }
}
