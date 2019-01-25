using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Linq;

namespace SPOMig
{
    /// <summary>
    /// This class is used to interact with SharePoint Online
    /// </summary>
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
            this.hashColumn = ConfigurationManager.AppSettings["HashColumn"];
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
        /// <param name="docLib">Name of the library to set</param>
        /// <returns>The list object including the list RootFolder for further processing</returns>
        public List setLibraryReadyForPRocessing(string docLib)
        {
            //We enable Folder creation for the SharePoint Online library
            List list = Context.Web.Lists.GetByTitle(docLib);
            list.EnableFolderCreation = true;
            list.Update();
            Context.Load(list, l => l.Title, l => l.RootFolder);
            Context.ExecuteQuery();

            //We try to retrieve the hashField
            try
            {
                Field hashField = list.Fields.GetByInternalNameOrTitle(this.hashColumn);
                Context.Load(hashField);
                Context.ExecuteQuery();
            }
            catch (ServerException ex)
            {
                //If we cannot retrieve the hashfield, we create it
                if (ex.Message.EndsWith("deleted by another user.") || ex.Message.Contains("Invalid field name"))
                {
                    string schemaTextField = $"<Field Type='Text' Name='{this.hashColumn}' StaticName='{this.hashColumn}' DisplayName='{this.hashColumn}' />";
                    Field simpleTextField = list.Fields.AddFieldAsXml(schemaTextField, false, AddFieldOptions.AddFieldInternalNameHint);
                    Context.ExecuteQuery();
                }
                else
                {
                    throw ex;
                }
            }
            return list;
        }

        /// <summary>
        /// Enable the folder creation in the SharePoint Online library and ensure the hash column is present
        /// </summary>
        /// <param name="docLib">Name of the library to set</param>
        /// <returns>The list object including the list RootFolder for further processing</returns>
        public bool cleanLibraryFromProcessing(string docLib)
        {
            //We enable Folder creation for the SharePoint Online library
            List list = Context.Web.Lists.GetByTitle(docLib);
            Context.Load(list.RootFolder);
            Context.ExecuteQuery();

            //We try to retrieve the hashField
            try
            {
                Field hashField = list.Fields.GetByInternalNameOrTitle(this.hashColumn);
                Context.Load(hashField);
                hashField.DeleteObject();
                list.Update();
                Context.ExecuteQuery();
                return true;
            }
            catch (ServerException ex)
            {
                //If we cannot retrieve the hashfield, job is done
                if (ex.Message.EndsWith("deleted by another user.") || ex.Message.EndsWith("invalid fieldname"))
                {
                    return true;
                }
                else
                {
                    return false;
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Retrieve all listitems in a library
        /// </summary>
        /// <returns></returns>
        public List<ListItem> GetAllDocumentsInaLibrary(string libName)
        {
            List<ListItem> items = new List<ListItem>();
            ClientContext ctx = this.Context;
            //ctx.Credentials = Your Credentials
            ctx.Load(ctx.Web, a => a.Lists);
            ctx.ExecuteQuery();

            List list = ctx.Web.Lists.GetByTitle("Documents");
            ListItemCollectionPosition position = null;
            // Page Size: 100
            int rowLimit = 100;
            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
            <Query>
                <OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy>
            </Query>
            <ViewFields>
            <FieldRef Name='Title'/><FieldRef Name='Modified' /><FieldRef Name='Editor' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' /><FieldRef Name='" + this.hashColumn + "' /></ViewFields><RowLimit Paged='TRUE'>" + rowLimit + "</RowLimit></View>";
            do
            {
                ListItemCollection listItems = null;
                camlQuery.ListItemCollectionPosition = position;
                listItems = list.GetItems(camlQuery);
                ctx.Load(listItems);
                ctx.ExecuteQuery();
                position = listItems.ListItemCollectionPosition;
                items.AddRange(listItems.ToList());
            }
            while (position != null);
            
            return items;
        }

        /// <summary>
        /// Copy folder to a SharePoint Online Site library
        /// </summary>
        /// <param name="folder">Folder to copy</param>
        /// <param name="list">List to copy folder to</param>
        /// <param name="localPath">Local Path selected by user</param>
        /// <returns>CopyStatus - the result from processing</returns>
        public CopyStatus copyFolderToSPO(DirectoryInfo folder, List list, string localPath, List<ListItem> onlineListItem)
        {
            //We instanciate the CopyStatus object to return
            CopyStatus copyStat = new CopyStatus
            {
                Name = folder.Name,
                Type = CopyStatus.ItemType.Folder,
                Path = folder.FullName.Remove(0, localPath.Length)
            };

            //We retrieve the normalized Urls (ItemNormalized path and ServerRelativeUrl)
            ItemURLs folderUrls = formatUrl(folder, list, localPath);

            //We stop processing if we detect the RootFolder (user selected path)
            if (folderUrls.ItemNormalizedPath == "") return null;

            //If the folder does not exist we create it
            if (checkFolderExist(folderUrls.ServerRelativeUrl, onlineListItem) == false)
            {

                /*
                //We create the folder ListITemCreationInformation
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderUrls.ItemNormalizedPath;

                //We create the folder
                ListItem newItem = list.AddItem(itemCreateInfo);
                
                //We update the folder metadata

                

                Context.ExecuteQuery();

                //We update the CopyStatus accordingly
                copyStat.Status = CopyStatus.ItemStatus.Created;
                copyStat.Comment = "Folder not found online - created";

                return copyStat;

    */

                

                
                var rootFolder = list.RootFolder;
                Context.Load(rootFolder);
                Context.ExecuteQuery();
                var myFolder = rootFolder.Folders.Add(folderUrls.ServerRelativeUrl);
                Context.ExecuteQuery();

                //We update metadate
                ListItem listitemFolder = Context.Web.GetListItem(folderUrls.ServerRelativeUrl);
                listitemFolder["Created"] = folder.CreationTimeUtc;
                listitemFolder["Modified"] = folder.CreationTimeUtc;
                listitemFolder.Update();
                Context.ExecuteQuery();

                //We update the CopyStatus accordingly
                copyStat.Status = CopyStatus.ItemStatus.Created;
                copyStat.Comment = "Folder not found online - created";

                return copyStat;

            }
            //The folder allready exists
            else
            {
                //We update the CopyStatus accordingly
                copyStat.Status = CopyStatus.ItemStatus.Skiped;
                copyStat.Comment = "Folder found online - skiped";

                return copyStat;
            }
        }

        /// <summary>
        /// Logic to copy file using file.add or StartUpload depending on file size
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="libraryName"></param>
        /// <param name="fileName"></param>
        /// <param name="fileChunkSizeInMB"></param>
        public void UploadFileSlicePerSlice(ClientContext ctx, string libraryName, string fileName, string itemNormalizedPath, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the folder to upload into. 
            List docs = ctx.Web.Lists.GetByTitle(libraryName);
            ctx.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // We create the file object
            Microsoft.SharePoint.Client.File uploadFile;

            // We calculate block size in bytes
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            // We retrieve the size of the file
            long fileSize = new FileInfo(fileName).Length;

            //If local file size < block size
            if (fileSize <= blockSize)
            {
                // We use File.add method to upload
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = itemNormalizedPath;
                    fileInfo.Overwrite = true;
                    uploadFile = docs.RootFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                }
            }
            else
            {
                // We use the large file method
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // We read the local file by block 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // We check if we read the last block 
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            // We check if we read the first block 
                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // We add an empty file
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = itemNormalizedPath;
                                    fileInfo.Overwrite = true;
                                    uploadFile = docs.RootFolder.Files.Add(fileInfo);

                                    // We start upload by uploading the first block 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first block
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // We set fileoffset as the pointer where the next slice will be added
                                        fileoffset = bytesUploaded.Value;
                                    }
                                    first = false;
                                }
                            }
                            else
                            {
                                // We get a reference to our file
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(itemNormalizedPath);

                                // We check if it is the last block
                                if (last)
                                {
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // We end the upload by calling FinishUpload
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                    }
                                }
                                else // We continue the upload
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next block.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } 
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// This method provide the logic to compare local and online file to choose either to copy or not
        /// </summary>
        /// <param name="file">File to copy</param>
        /// <param name="list">The Targeted SharePoint Online Library</param>
        /// <param name="localPath">Local Path selected by user</param>
        /// <returns>CopyStatus - The result of processing</returns>
        public CopyStatus copyFileToSPO(FileInfo file, List list, string localPath, List<ListItem> onlineListItem)
        {
            //We instanciate the CopyStatus object to return feedback from processing
            CopyStatus copystat = new CopyStatus()
            {
                Name = file.Name,
                Type = CopyStatus.ItemType.File,
                Path = file.FullName.Remove(0, localPath.Length)
            };

            //We set the variable for the using statement bellow
            ItemURLs fileUrls;
            DateTime created;
            DateTime modified;
            string localFileLength;
            string localFileHash;
            string siteURL;
            string itemUrl;
            OnlineFileStatus targetFileStat;

            //using the FileStream to dispose when processing is over
            using (FileStream fileStream = new FileStream(file.FullName, FileMode.Open))
            {
                //We retrieve the normalized Urls (ItemNormalized path and ServerRelativeUrl)
                fileUrls = formatUrl(file, list, localPath);

                //We retrieve the local file metadata
                created = file.CreationTimeUtc;
                modified = file.LastWriteTimeUtc;
                localFileLength = file.Length.ToString();
                localFileHash = FileLogic.hashFromLocal(fileStream);

                //We retrieve the ListItem URL to check if it exists on the SharePoint Online library
                siteURL = getSiteURL();
                itemUrl = siteURL + "/" + list.RootFolder.Name + "/" + fileUrls.ItemNormalizedPath;

                //We retrive the target file length (does not exist == 0)
                targetFileStat = checkListItemExist(fileUrls.ServerRelativeUrl, onlineListItem);
            }

            //If the target item does not exist, we copy the file and set metadata
            if (targetFileStat.FileFound == OnlineFileStatus.FileStatus.NotFound)
            {
                //We copy the file and set metadata
                //Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, fileUrls.ServerRelativeUrl, fileStream, false);
                UploadFileSlicePerSlice(Context, list.Title, file.FullName, fileUrls.ServerRelativeUrl);
                setUploadedFileMetadata(fileUrls.ServerRelativeUrl, created, modified, localFileHash);

                //We update the CopyStatus accordingly
                copystat.Status =  CopyStatus.ItemStatus.Uploaded;
                copystat.Comment = "File not found online - Uploaded";

                return copystat;
            }
            //If target item has no hash, we compare lenght to check if copy is necessary
            else if (targetFileStat.HashFound == OnlineFileStatus.HashStatus.NotFound)
            {
                //We retrive the target file length
                string targetFileLength = getFileLenght(fileUrls.ServerRelativeUrl);
                    
                //Local and Online Files are the same length, se we do not overwrite
                if (localFileLength == targetFileLength)
                {
                    //We update metadata
                    setUploadedFileMetadata(fileUrls.ServerRelativeUrl, created, modified, localFileHash);

                    //We update the CopyStatus accordingly
                    copystat.Status = CopyStatus.ItemStatus.Skiped;
                    copystat.Comment = "File found online but not hash - files are the same length so we do not overwrite the online file";

                    return copystat;
                }
                else//If the file are different length, we overwrite the file and set metadata
                {
                    //We copy the file and set metadata
                    //Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, fileUrls.ServerRelativeUrl, fileStream, true);
                    UploadFileSlicePerSlice(Context, list.Title, file.FullName, fileUrls.ServerRelativeUrl);
                    setUploadedFileMetadata(fileUrls.ServerRelativeUrl, created, modified, localFileHash);

                    //We update the CopyStatus accordingly
                    copystat.Status = CopyStatus.ItemStatus.Overwrited;
                    copystat.Comment = "File found online but not hash - files are not the same length so we overwrite the online file";

                    return copystat;
                }
            }
            else
            {
                //If files are the same Hash, we do not overwrite
                if (localFileHash == targetFileStat.Hash)
                {
                    //We update the CopyStatus accordingly
                    copystat.Status = CopyStatus.ItemStatus.Skiped;
                    copystat.Comment = "File found online - files are the same hash so we do not overwrite the online file";

                    return copystat;
                }
                else //The file are not the same hash, we overwrite it and set metadata
                {
                    //We copy the file and set metadata
                    //Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, fileUrls.ServerRelativeUrl, fileStream, true);
                    UploadFileSlicePerSlice(Context, list.Title, file.FullName, fileUrls.ServerRelativeUrl);
                    setUploadedFileMetadata(fileUrls.ServerRelativeUrl, created, modified, localFileHash);

                    //We update the CopyStatus accordingly
                    copystat.Status = CopyStatus.ItemStatus.Overwrited;
                    copystat.Comment = "File found online - files are not the same hash so we overwrite the online file";

                    return copystat;
                }
                
            }
        }

        /// <summary>
        /// Retrieve the necessary URLs to process folders related operations
        /// </summary>
        /// <param name="folder">DirectoryInfo of the folder to process</param>
        /// <param name="list">The Targeted SharePoint Online Library</param>
        /// <param name="localPath">Local Path selected by user - Used ton normalize the urls</param>
        /// <returns>ItemURLs - Object containing the Normalized path and the ServerRelativeURL</returns>
        private ItemURLs formatUrl(DirectoryInfo folder, List list, string localPath)
        {
            ItemURLs folderUrls = new ItemURLs();

            //We retrieve the folder Normalized path
            string localPathNormalized = localPath.Replace("/", "\\");
            string folderPath = folder.FullName.Replace("/", "\\");
            string folderPathNormalized = folderPath.Replace(localPathNormalized, "");
            string folderPathNormalizedFinal = folderPathNormalized.Replace("\\", "/");
            folderUrls.ItemNormalizedPath = folderPathNormalizedFinal;

            //We retrieve the folder ServerRelativeUrl
            string libURL = list.RootFolder.ServerRelativeUrl.ToString();
            string serverRelativeURL = libURL + "/" + folderPathNormalizedFinal;
            folderUrls.ServerRelativeUrl = serverRelativeURL;

            return folderUrls;
        }

        /// <summary>
        /// Retrieve the necessary URLs to process file related operations
        /// </summary>
        /// <param name="folder">FileInfo of the file to process</param>
        /// <param name="list">The Targeted SharePoint Online Library</param>
        /// <param name="localPath">Local Path selected by user - Used ton normalize the urls</param>
        /// <returns>ItemURLs - Object containing the Normalized path and the ServerRelativeURL</returns>
        private ItemURLs formatUrl(FileInfo file, List list, string localPath)
        {
            ItemURLs itemUrls = new ItemURLs();

            //We retrieve the ItemNormalized path
            string libURL = list.RootFolder.ServerRelativeUrl.ToString();
            string localPathNormalized = localPath.Replace("/", "\\");
            string filePath = file.FullName.Replace("/", "\\");
            string fileNormalizedPath = filePath.Replace(localPathNormalized, "");
            string fileNormalizedPathfinal = fileNormalizedPath.Replace("\\", "/");
            itemUrls.ItemNormalizedPath = fileNormalizedPathfinal;

            //We contruct the item ServerRelativeUrl
            string serverRelativeURL = libURL + "/" + fileNormalizedPathfinal;
            itemUrls.ServerRelativeUrl = serverRelativeURL;

            return itemUrls;

        }

        /// <summary>
        /// Update the Created, Modified and FileHash field of a ListItem in a SharePoint Online Library 
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
        private bool checkFolderExist (string itemPath, List<ListItem> onlineListItem)
        {
            foreach (var item in onlineListItem)
            {
                if ((string)item["FileRef"] == itemPath )
                {
                    return true;
                }
            }
            return false;
        }
        
        /// <summary>
        /// Verify if the item allready exist in the SharePoint Online library and retrieve the fileHash column value
        /// </summary>
        /// <param name="itemPath"></param>
        /// <returns>OnlineFileStatus (FileFound?,HashFound?,HashValue)</returns>
        private OnlineFileStatus checkListItemExist(string itemPath, List<ListItem> onlineListItem)
        {
            //We instanciate the OnlineFileStatus object
            OnlineFileStatus status = new OnlineFileStatus();

            foreach (var item in onlineListItem) 
            {
                // If we find the listitem
                if ((string)item["FileRef"] == itemPath)
                {
                    //At this point we found the file so we update the OnlineFileStatus accordingly
                    status.FileFound = OnlineFileStatus.FileStatus.Found;

                    //We try to retrieve the value from the hashColumn
                    string fileHash = item[this.hashColumn].ToString();
                    //At this point we found the HashColumn so we update the OnlineFileStatus accordingly
                    status.HashFound = OnlineFileStatus.HashStatus.Found;
                    status.Hash = fileHash;

                    return status;
                }
            }
            //We update the OnlineFileStatus accordingly
            status.FileFound = OnlineFileStatus.FileStatus.NotFound;
            status.HashFound = OnlineFileStatus.HashStatus.NotFound;
            status.Hash = null;

            return status;
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
