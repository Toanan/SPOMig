using System.Windows;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.Configuration;

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
        #endregion

        #region Ctor

        public ClientContext Context { get; set; }
        public BulkWindow()
        {
            InitializeComponent();
        }

        #endregion


        private void Btn_Connect_Click(object sender, RoutedEventArgs e)
        {

            string ID = ConfigurationManager.AppSettings["AppID"];
            string Sec = ConfigurationManager.AppSettings["Secret"];
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

        private void BtnHome_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            Startup mw = new Startup();
            mw.Show();
            this.Close();
        }
    }
}
