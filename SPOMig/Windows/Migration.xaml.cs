using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using System.IO;

namespace SPOMig
{
    /// <summary>
    /// Logique d'interaction pour Migration.xaml
    /// </summary>
    public partial class Migration : Window
    {
        #region Props
        public string LocalPath { get; set; }
        public ClientContext Context { get; set; }
        #endregion

        #region Ctor
        public Migration(ListCollection ListCollection, ClientContext ctx)
        {
            InitializeComponent();
            this.Context = ctx;
            // Retrive only document library from the ListCollection passed
            foreach (List list in ListCollection)
            {
                if (list.BaseTemplate == 101)
                    Cb_doclib.Items.Add(list.Title);
            }
        }
        #endregion

        #region EventHandler
        /// <summary>
        /// Button Copy OnClick event :
        /// Input check then BackgroundWorker launch
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Copy_Click(object sender, RoutedEventArgs e)
        {
            this.LocalPath = @Tb_LocalPath.Text;
            if (LocalPath == "")
            {
                MessageBox.Show("Please fill the local path field");
                return;
            }
            if (!Directory.Exists(LocalPath))
            {
                MessageBox.Show("Cannot find local path, please double check");
                return;
            }
            if (Cb_doclib.SelectedItem == null)
            {
                MessageBox.Show("Please select a document library");
                return;
            }

            FileLogic FileLogic = new FileLogic(LocalPath);
            List<FileInfo> files = FileLogic.getFiles();
            SPOLogic ctx = new SPOLogic(Context);
            ctx.copyFileToSPO(Cb_doclib.Text, files);
        }
        #endregion

    }
}
