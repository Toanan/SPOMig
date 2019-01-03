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
        public string SiteUrl { get; set; }
        public string UserName { get; set; }
        public SecureString PassWord { get; set; }
        public ListCollection Lists { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Lb_Top.Content = "First connect to a SharePoint Online Site";
            bw.DoWork += bw_Dowork;
            bw.ProgressChanged += bw_ProgressChanged;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Hide();
            Migration mig = new Migration(Lists);
            mig.Show();
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
        }

        private void bw_Dowork(object sender, DoWorkEventArgs e)
        {
            var ctx = new ClientContext(SiteUrl);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(UserName, PassWord);
            ctx.Credentials = credentials;
            SPOLogic Context = new SPOLogic(ctx);
            ListCollection lists = Context.getLists();
            this.Lists = lists;
        }

        private void Btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.SiteUrl = Tb_SPOSite.Text;
                this.UserName = Tb_UserName.Text;
                this.PassWord = Pb_PassWord.SecurePassword;
            }
            catch
            {
                MessageBox.Show("renseignez les champs");
            }
            bw.RunWorkerAsync();

            
        }
    }
}
