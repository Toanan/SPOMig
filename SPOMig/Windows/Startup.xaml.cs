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
    /// Logique d'interaction pour Startup.xaml
    /// </summary>
    public partial class Startup : Window
    {
        public Startup()
        {
            InitializeComponent();

            // We check for the cfg xml file to exist
            FileLogic.ensureConfigFileExists();
        }

        private void Btn_GranularScenario_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            MainWindow mw = new MainWindow();
            mw.Show();
            this.Close();
        }

        private void Btn_BulkScenario_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();

            string appID = FileLogic.getXMLSettings("appID");
            string appSecret = FileLogic.getXMLSettings("appSecret");

            //We check if appId and secret are configured
            if (string.IsNullOrWhiteSpace(appID) || string.IsNullOrWhiteSpace(appSecret))
            {
                this.Hide();
                AppOnlyConfig appConfig = new AppOnlyConfig();
                appConfig.Show();
                this.Close();
            }
            else
            {
                BulkWindow bw = new BulkWindow();
                bw.Show();
                this.Close();
            }
        }
    }
}
