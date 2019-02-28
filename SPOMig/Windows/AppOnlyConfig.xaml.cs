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
using System.Configuration;

namespace SPOMig.Windows
{
    /// <summary>
    /// Logique d'interaction pour AppOnlyConfig.xaml
    /// </summary>
    public partial class AppOnlyConfig : Window
    {
        public AppOnlyConfig()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Btn_AppOnlyCfg event : Onclick() Apply AppOnly configuration to the app.config file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_AppOnlyCfg_Click(object sender, RoutedEventArgs e)
        {
            FileLogic.setXMLSettingsAttribute("appID", Tb_AppID.Text);
            FileLogic.setXMLSettingsAttribute("appSecret", Tb_AppSecret.Text);
            this.Hide();
            BulkWindow bulkWindow = new BulkWindow();
            bulkWindow.Show();
            this.Close();
        }
    }
}
