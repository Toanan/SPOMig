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

        public void copyFileToSPO(FileInfo file, List list, string localPath)
        {
            using (FileStream fileStream = new FileStream(file.FullName, FileMode.Open))
            {
                string libURL = list.RootFolder.ServerRelativeUrl.ToString();
                string localPathNormalized = localPath.Replace("/", "\\");
                string filePath = file.FullName.Replace("/", "\\");
                string fileNormalizedPath = filePath.Replace(localPathNormalized, "");
                string fileNormalizedPathfinal = fileNormalizedPath.Replace("\\", "/");
                string serverRelativeURL = libURL + "/" + fileNormalizedPathfinal;

 

                Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, true);
            }
        }

        public void copyFolderToSPO (DirectoryInfo folder, List list, string localPath)
        {
            string localPathNormalized = localPath.Replace("/", "\\");
            string folderPath = folder.FullName.Replace("/", "\\");
            string folderPathNormalized = folderPath.Replace(localPathNormalized, "");
            string folderPathNormalizedFinal = folderPathNormalized.Replace("\\", "/");
            if (folderPathNormalizedFinal == "") return;

            if (checkItemExist(folderPathNormalizedFinal) == false)
            {
                //To create the folder
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderPathNormalizedFinal;

                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem["Title"] = folderPathNormalizedFinal;
                newItem.Update();
                Context.ExecuteQuery();
            }   
        }

        private bool checkItemExist (string itemPath)
        {
            try
            {
                ListItem folderToCheck = Context.Web.GetListItem(itemPath);
                Context.Load(folderToCheck);
                Context.ExecuteQuery();
                return true;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return false;
                }
                throw;
            }
        }
        #endregion

    }
}
