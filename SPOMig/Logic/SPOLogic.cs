using System;
using System.IO;
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

        /// <summary>
        /// Copy File to a SharePoint Online library
        /// </summary>
        /// <param name="file">File to copy</param>
        /// <param name="list">List to copy file to</param>
        /// <param name="localPath">Local Path selected by user - To normalize folder path in the library</param>
        public bool copyFileToSPO(FileInfo file, List list, string localPath)
        {
            //using the FileStream to dispose when computing is over
            using (FileStream fileStream = new FileStream(file.FullName, FileMode.Open))
            {
                #region URL formating
                //We retrieve the library serverRelativeUrl, localFilePath to compute the listItem Full Url
                string libURL = list.RootFolder.ServerRelativeUrl.ToString();
                string localPathNormalized = localPath.Replace("/", "\\");
                string filePath = file.FullName.Replace("/", "\\");
                string fileNormalizedPath = filePath.Replace(localPathNormalized, "");
                string fileNormalizedPathfinal = fileNormalizedPath.Replace("\\", "/");
                string serverRelativeURL = libURL + "/" + fileNormalizedPathfinal;
                #endregion

                //We retrive the local file metadata
                DateTime created = file.CreationTimeUtc;
                DateTime modified = file.LastWriteTimeUtc;

                //We retrive the ListItem URL to check if it exists on the SharePoint Online library
                string siteURL = getSiteURL();
                string itemUrl = siteURL + "/" + list.RootFolder.Name + "/" + fileNormalizedPathfinal;

                //We retrive the target file length (does not exist == 0)
                long targetLenght = checkItemExist(itemUrl);

                //If the item doesn't exist => we copy the file
                if (targetLenght == 0)
                {
                    //We copy the file
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, false);
                    //And set the metadata
                    setUploadedFileMetadata(serverRelativeURL, created, modified);

                    /* Hash Commented as we user item.lenght to compare files
                    //We compute the hash
                    string currentHash = FileLogic.hashFromLocal(fileStream);
                    //We update the listitem providing the hash to the custom column
                    setUploadedFileHash(serverRelativeURL, currentHash);
                    */

                    return true;
                }
                else //File allready exist => compare file sizes
                {
                    //We retrive the local file length
                    long localLenght = file.Length;

                    //Check if the file are the same length
                    if (localLenght == targetLenght)
                    {
                        //Yes, do nothing
                        return false;
                    }
                    else //The file has changed, so we overwrite it and set metadata
                    {
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, true);
                        setUploadedFileMetadata(serverRelativeURL, created, modified);
                        return true;
                    }
                }
            }
        }

        /// <summary>
        /// Update the Created and modified field using local file metadata
        /// </summary>
        /// <param name="serverRelativeURL"></param>
        /// <param name="created"></param>
        /// <param name="modified"></param>
        private void setUploadedFileMetadata(string serverRelativeURL, DateTime created, DateTime modified)
        {
            ListItem currentOnlinefile = Context.Web.GetListItem(serverRelativeURL);
            currentOnlinefile["Created"] = created;
            currentOnlinefile["Modified"] = modified;
            currentOnlinefile.Update();
            Context.ExecuteQuery();
        }

        /// <summary>
        /// Update a file to add the hash value to a custom column
        /// </summary>
        /// <param name="serverRelativeURL"></param>
        /// <param name="hash"></param>
        private void setUploadedFileHash(string serverRelativeURL, string hash)
        {
            ListItem currentOnlinefile = Context.Web.GetListItem(serverRelativeURL);
            currentOnlinefile["test"] = hash;
            currentOnlinefile.Update();
            Context.ExecuteQuery();
        }

        /// <summary>
        /// Retrive the SharePointOnline site URL
        /// </summary>
        /// <returns>SharePointSite URL</returns>
        private string getSiteURL()
        {
            Web site = Context.Web;
            Context.Load(site, s => s.Url);
            Context.ExecuteQuery();
            return site.Url;
        }

        /// <summary>
        /// Copy folder to a SharePoint Online Site library
        /// </summary>
        /// <param name="folder">Folder to copy</param>
        /// <param name="list">List to copy folder to</param>
        /// <param name="localPath">Local Path selected by user - To normalize folder path in the library</param>
        public void copyFolderToSPO (DirectoryInfo folder, List list, string localPath)
        {
            string localPathNormalized = localPath.Replace("/", "\\");
            string folderPath = folder.FullName.Replace("/", "\\");
            string folderPathNormalized = folderPath.Replace(localPathNormalized, "");
            string folderPathNormalizedFinal = folderPathNormalized.Replace("\\", "/");
            if (folderPathNormalizedFinal == "") return;

            if (checkFolderExist(folderPathNormalizedFinal) == false)
            {
                //To create the folder
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderPathNormalizedFinal;

                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem["Title"] = folderPathNormalizedFinal;
                newItem["Created"] = folder.CreationTimeUtc;
                newItem["Modified"] = folder.CreationTimeUtc;
                newItem.Update();
                Context.ExecuteQuery();
            }   
        }

        /// <summary>
        /// Verify if the folder allready exist in the SharePoint Online library
        /// </summary>
        /// <param name="itemPath"></param>
        /// <returns>Yes or no</returns>
        private bool checkFolderExist (string itemPath)
        {
            try
            {
                ListItem itemtoCheck = Context.Web.GetListItem(itemPath);
                Context.Load(itemtoCheck);
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

        /// <summary>
        /// Verify if the item allready exist in the SharePoint Online library
        /// </summary>
        /// <param name="itemPath"></param>
        /// <returns>File lenght</returns>
        private long checkItemExist (string itemPath)
        {
            try
            {
                Microsoft.SharePoint.Client.File file = Context.Web.GetFileByUrl(itemPath);
                Context.Load(file);
                Context.ExecuteQuery();
                return file.Length;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return 0;
                }
                throw;
            }
        }

        
        #endregion

    }
}
