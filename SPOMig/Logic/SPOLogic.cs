using System;
using System.IO;
using Microsoft.SharePoint.Client;

namespace SPOMig
{
    class SPOLogic
    {
        #region Props
        public ClientContext Context { get; set; }

        public string hashColumn { get; set; }
        #endregion

        #region Ctor
        public SPOLogic(ClientContext ctx)
        {
            this.Context = ctx;
            this.hashColumn = "FileHash";
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
        /// Enable the folder creation in the SharePoint Online library and ensure the hash column is present
        /// </summary>
        /// <param name="docLib"></param>
        /// <returns></returns>
        public List setLibraryReadyForPRocessing(string docLib)
        {
            //We enable Folder creation for the SharePoint Online library
            List list = Context.Web.Lists.GetByTitle(docLib);
            list.EnableFolderCreation = true;
            list.Update();
            Context.Load(list.RootFolder);
            Context.ExecuteQuery();
            try
            {
                Field hashField = list.Fields.GetByInternalNameOrTitle(this.hashColumn);
                Context.Load(hashField);
                Context.ExecuteQuery();
            }

            catch (ServerException ex)
            {
                if (ex.Message.EndsWith("deleted by another user."))
                {
                    string schemaTextField = $"<Field Type='Text' Name='{this.hashColumn}' StaticName='{this.hashColumn}' DisplayName='{this.hashColumn}' />";
                    Field simpleTextField = list.Fields.AddFieldAsXml(schemaTextField, false, AddFieldOptions.AddFieldInternalNameHint);
                    Context.ExecuteQuery();
                }
                else
                {
                    throw;
                }
            }
            return list;
        }

        /// <summary>
        /// Copy folder to a SharePoint Online Site library
        /// </summary>
        /// <param name="folder">Folder to copy</param>
        /// <param name="list">List to copy folder to</param>
        /// <param name="localPath">Local Path selected by user - To normalize folder path in the library</param>
        public CopyStatus copyFolderToSPO(DirectoryInfo folder, List list, string localPath)
        {
            #region URL formating
            string localPathNormalized = localPath.Replace("/", "\\");
            string folderPath = folder.FullName.Replace("/", "\\");
            string folderPathNormalized = folderPath.Replace(localPathNormalized, "");
            string folderPathNormalizedFinal = folderPathNormalized.Replace("\\", "/");
            if (folderPathNormalizedFinal == "") return null;

            string libURL = list.RootFolder.ServerRelativeUrl.ToString();
            string serverRelativeURL = libURL + "/" + folderPathNormalizedFinal;
            #endregion

            CopyStatus copyStat = new CopyStatus
            {
                Name = folder.Name,
                Type = "Folder",
                Path = folder.FullName.Remove(0, localPath.Length)
            };

            //If the folder does not exist we create it
            if (checkFolderExist(serverRelativeURL) == false)
            {
                //We create the folder
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderPathNormalizedFinal;

                ListItem newItem = list.AddItem(itemCreateInfo);
                newItem["Title"] = folderPathNormalizedFinal;
                newItem["Created"] = folder.CreationTimeUtc;
                newItem["Modified"] = folder.CreationTimeUtc;
                newItem.Update();
                Context.ExecuteQuery();

                copyStat.Status = "Created";
                copyStat.Comment = "Folder not found online - created";

                return copyStat;


            }
            else
            {
                copyStat.Status = "Skiped";
                copyStat.Comment = "Folder found online - skiped";

                return copyStat;
            }
        }

        /// <summary>
        /// Copy File to a SharePoint Online library
        /// </summary>
        /// <param name="file">File to copy</param>
        /// <param name="list">List to copy file to</param>
        /// <param name="localPath">Local Path selected by user - To normalize folder path in the library</param>
        public CopyStatus copyFileToSPO(FileInfo file, List list, string localPath)
        {
            //We instanciate the CopyStatus object to return
            CopyStatus copystat = new CopyStatus()
            {
                Name = file.Name,
                Type = "File",
                Path = file.FullName.Remove(0, localPath.Length)
            };

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
                string localFileLength = file.Length.ToString();
                string localFileHash = FileLogic.hashFromLocal(fileStream);

                //We retrive the ListItem URL to check if it exists on the SharePoint Online library
                string siteURL = getSiteURL();
                string itemUrl = siteURL + "/" + list.RootFolder.Name + "/" + fileNormalizedPathfinal;

                //We retrive the target file length (does not exist == 0)
                string targetFileHash = checkItemExist(itemUrl);

                //If the target item doesn't exist => we copy the file and set metadata
                if (targetFileHash == "notFound")
                {
                    //We copy the file
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, false);
                    //And set the metadata
                    setUploadedFileMetadata(serverRelativeURL, created, modified, localFileHash);
                    copystat.Status = "Uploaded";
                    copystat.Comment = "File not found online - Uploaded";

                    return copystat;
                }
                //If target item has no hash => we compare lenght to check if copy is necessary
                else if (targetFileHash == "noHash")
                {
                    string targetFileLength = getFileLenght(serverRelativeURL);
                    //Same length => no copy
                    if (localFileLength == targetFileLength)
                    {
                        copystat.Status = "Skiped";
                        copystat.Comment = "File found online but not hash - files are the same length so we do not upload";
                        return copystat;
                    }
                    //Different length => wet overwrite the file and set metadata
                    else
                    {
                        //We copy the file
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, true);
                        //And set the metadata
                        setUploadedFileMetadata(serverRelativeURL, created, modified, localFileHash);

                        copystat.Status = "Overwrited";
                        copystat.Comment = "File found online but not hash - files are not the same length so we overwrite the online file";

                        return copystat;
                    }
                }
                else
                {
                    //Check if the file are the same hash
                    if (localFileHash == targetFileHash)
                    {
                        //Yes, do nothing
                        copystat.Status = "Skiped";
                        copystat.Comment = "File found online - files are the same hash so we do not upload";
                        return copystat;
                    }
                    else //The file has changed, so we overwrite it and set metadata
                    {
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, true);
                        setUploadedFileMetadata(serverRelativeURL, created, modified, localFileHash);

                        copystat.Status = "Overwrite";
                        copystat.Comment = "File found online - files are not the same hash so we overwrite the online file";

                        return copystat;
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
        private void setUploadedFileMetadata(string serverRelativeURL, DateTime created, DateTime modified, string hash)
        {
            ListItem currentOnlinefile = Context.Web.GetListItem(serverRelativeURL);
            currentOnlinefile["Created"] = created;
            currentOnlinefile["Modified"] = modified;
            currentOnlinefile[this.hashColumn] = hash;
            currentOnlinefile.Update();
            Context.ExecuteQuery();
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
        private string checkItemExist(string itemPath)
        {
            try
            {
                ListItem file = Context.Web.GetListItem(itemPath);
                Context.Load(file);
                Context.ExecuteQuery();
                string fileHash = file[this.hashColumn].ToString();
                return fileHash;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return "notFound";
                }
                throw;
            }
            catch (Exception ex)
            {
                if (ex.Message == "Object reference not set to an instance of an object.")
                {
                    return "noHash";
                }
                throw;
            }
        }

        /// <summary>
        /// Retrieve SharePoint Online file lenght
        /// </summary>
        /// <param name="itemPath"></param>
        /// <returns>File lenght as string</returns>
        private string getFileLenght (string itemPath)
        {
            Microsoft.SharePoint.Client.File file = Context.Web.GetFileByUrl(itemPath);
            Context.Load(file, f => f.Length);
            Context.ExecuteQuery();
            return file.Length.ToString();
        }

        #endregion

    }
}
