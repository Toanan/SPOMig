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
using System.IO;

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

        #region EvenHandler

        /// <summary>
        /// Opens the user login windows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_GranularScenario_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            MainWindow mw = new MainWindow();
            mw.Show();
            this.Close();
        }

        /// <summary>
        /// Opens the AppOnly login window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Extract all items in a path to create a CSV file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Extract_Click(object sender, RoutedEventArgs e)
        {
            string localPath = @Tb_LocalPath.Text;

            #region UserInput checks
            if (localPath == "")
            {
                MessageBox.Show("Please fill the local path field");
                return;
            }
            if (!Directory.Exists(localPath))
            {
                MessageBox.Show("Cannot find local path, please double check");
                return;
            }
            /*
            if (Cb_doclib.SelectedItem == null)
            {
                MessageBox.Show("Please select a document library");
                return;
            }*/
            #endregion

            // We retrieve the source folder information to compute the extract file path and create it
            DirectoryInfo sourceFolder = new DirectoryInfo(localPath);
            Reporting extractCSV = new Reporting(sourceFolder.Name);

            //We retrive the list of DirectoryInfo and FileInfo
            List<FileInfo> files = FileLogic.getFiles(localPath);
            List<DirectoryInfo> folders = FileLogic.getSourceFolders(localPath);

            // We loop every files then folders to create a file mapping and extrat metadata to a csv file
            foreach (FileInfo file in files)
            {
                FileMapping filemap = new FileMapping
                {
                    Name = file.Name,
                    Path = file.FullName,
                    Modified = file.LastWriteTimeUtc,
                    Created = file.CreationTimeUtc,
                    Owner = File.GetAccessControl(file.FullName).GetOwner(typeof(System.Security.Principal.NTAccount)).ToString(),
                    ItemType = FileMapping.Type.File
                };
                extractCSV.writeExtract(filemap);
            }

            foreach (DirectoryInfo dir in folders)
            {
                FileMapping filemap = new FileMapping
                {
                    Name = dir.Name,
                    Path = dir.FullName,
                    Modified = dir.LastWriteTimeUtc,
                    Created = dir.CreationTimeUtc,
                    Owner = File.GetAccessControl(dir.FullName).GetOwner(typeof(System.Security.Principal.NTAccount)).ToString(),
                    ItemType = FileMapping.Type.Folder
                };
                extractCSV.writeExtract(filemap);
            }
            MessageBox.Show("Extraction done");
        }
        #endregion

    }
}
