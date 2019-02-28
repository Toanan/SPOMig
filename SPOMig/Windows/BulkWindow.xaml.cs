
using System.Windows;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Configuration;
using System;

namespace SPOMig.Windows
{

    /// <summary>
    /// Logique d'interaction pour BulkWindow.xaml
    /// </summary>
    public partial class BulkWindow : Window
    {
        #region Props

        public string AppID { get; set; }
        public string AppSecret { get; set; }
        public Configuration Config { get; set; }
        public ClientContext Context { get; set; }

        #endregion

        #region Ctor

        public BulkWindow(Configuration cfg)
        {
            InitializeComponent();
            this.Config = cfg;
        }

        #endregion


        private void Btn_Connect_Click(object sender, RoutedEventArgs e)
        {

            string ID = Config.AppSettings.Settings["AppID"].Value;
            string Sec = Config.AppSettings.Settings["Secret"].Value;
            int buff;
            Int32.TryParse(Config.AppSettings.Settings["Buffer"].Value, out buff);
            string site = Tb_CsvFile.Text;

            using (ClientContext ctx = new AuthenticationManager().GetAppOnlyAuthenticatedContext(site, ID, Sec))
            {
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                this.Context = ctx;
            }

            SPOLogic spol = new SPOLogic(Context);

            var items = spol.GetAllDocumentsInaLibrary("Documents");
            MessageBox.Show(items.Count.ToString());
        }

        /// <summary>
        /// Btn_Home OnClick event, open the startup window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnHome_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            Startup mw = new Startup();
            mw.Show();
            this.Close();
        }

        /// <summary>
        /// Btn_Config Onclick event, open the AppOnly configuration window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Config_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            AppOnlyConfig appConfig = new AppOnlyConfig();
            appConfig.Show();
            this.Close();
        }
    }
}
