using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
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
        #endregion

        #region Ctor
        public MainWindow()
        {
            InitializeComponent();
            Lb_Top.Content = "First connect to a SharePoint Online Site";  
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
            if (e.Cancelled == false)
            {
                this.Hide();
                Migration mig = new Migration(Lists, Context);
                mig.Show();
                bw.Dispose();
            }
            else
            {
                Btn_Connect.IsEnabled = true;
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
            try
            {
                using (var ctx = new ClientContext(SiteUrl))
                {
                    SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(UserName, PassWord);
                    ctx.Credentials = credentials;
                    this.Context = ctx;
                    SPOLogic Context = new SPOLogic(ctx);
                    ListCollection lists = Context.getLists();
                    this.Lists = lists;
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

            //BarckgroundWorker delegates
            bw.DoWork += bw_Dowork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;

            Btn_Connect.IsEnabled = false;

            this.SiteUrl = Tb_SPOSite.Text;
            this.UserName = Tb_UserName.Text;
            this.PassWord = Pb_PassWord.SecurePassword;

            bw.WorkerSupportsCancellation = true;
            bw.RunWorkerAsync();

        }

        #endregion

        #endregion

    }
}
