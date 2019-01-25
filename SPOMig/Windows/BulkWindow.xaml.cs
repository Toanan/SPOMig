using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.IO;
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

            foreach (var item in items)
            {
                System.IO.File.AppendAllText(@"c:/listallfile.txt", $"{item["Title"]},{item["FileRef"]},{item["FileLeafRef"]}");
            }



        }
    }
}
