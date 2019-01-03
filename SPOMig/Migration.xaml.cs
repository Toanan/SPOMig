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

namespace SPOMig
{
    /// <summary>
    /// Logique d'interaction pour Migration.xaml
    /// </summary>
    public partial class Migration : Window
    {

        public Migration(ListCollection ListCollection)
        {
            InitializeComponent();
            foreach (List list in ListCollection)
            {
                if (list.BaseTemplate == 101)
                    Cb_doclib.Items.Add(list.Title); 
            }
        }
    }
}
