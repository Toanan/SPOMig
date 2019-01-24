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

namespace SPOMig.Windows
{
    /// <summary>
    /// Logique d'interaction pour BulkWindow.xaml
    /// </summary>
    public partial class BulkWindow : Window
    {
        public ClientContext Context { get; set; }
        public BulkWindow()
        {
            InitializeComponent();
        }

        private void Btn_Connect_Click(object sender, RoutedEventArgs e)
        {
            string ID = Tb_AppId.Text;
            string Sec = TB_AppSecret.Text;
            string site = Tb_CsvFile.Text;

            AuthenticationManager am = new AuthenticationManager();

            using (ClientContext ctx = am.GetAppOnlyAuthenticatedContext(site, ID, Sec))
            {
                ClientContext context = ctx;
                ctx.ExecuteQuery();
                this.Context = context;
            }

            SPOLogic spol = new SPOLogic(Context);

            ListCollection lists = spol.getLists();
            MessageBox.Show(lists.Count.ToString());           
        }
    }
}
