using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SPOMig
{
    class SPOLogic
    {
        #region Props
        public ClientContext Context { get; set; }
        #endregion

        #region Ctor
        public SPOLogic(ClientContext ctx)
        {
            this.Context = ctx;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Connect to a SPO Site to retrieve the site lists
        /// </summary>
        /// <returns>ListCollection</returns>
        public ListCollection getLists()
        {
            ListCollection Libraries = Context.Web.Lists;
            Context.Load(Libraries);
            Context.ExecuteQuery();
            return Libraries;
        }

        public void copyFileToSPO(string libName, List<FileInfo> files)
        {
            List list = Context.Web.Lists.GetByTitle(libName);
            Context.Load(list.RootFolder);
            Context.ExecuteQuery();
            foreach (FileInfo file in files)
            {
                using (FileStream fileStream = new FileStream(file.FullName, FileMode.Open))
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, list.RootFolder.ServerRelativeUrl.ToString() + "/" + file.FullName.Split('\\')[1], fileStream, true);
                }
            }
        }
        #endregion

    }
}
