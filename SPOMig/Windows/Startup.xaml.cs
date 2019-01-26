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
            string appPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appName = "SPOMig";
            string appCfgPath = $"{appPath}/{appName}/SPOMig.cfg";
            ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
            configMap.ExeConfigFilename = appCfgPath;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);

            if (!System.IO.File.Exists(appCfgPath))
            {
                this.Hide();
                AppOnlyConfig appConfig = new AppOnlyConfig();
                appConfig.Show();
                this.Close();
            }
            else
            {
                BulkWindow bw = new BulkWindow(config);
                bw.Show();
                this.Close();
            }
        }
    }
}
