﻿using System;
using System.IO;
using Microsoft.SharePoint.Client;

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
        /// <param name="docLib">Name of the library to set</param>
        /// <returns>The list object including the list RootFolder for further processing</returns>
        public List setLibraryReadyForPRocessing(string docLib)
        {
            //We enable Folder creation for the SharePoint Online library
            List list = Context.Web.Lists.GetByTitle(docLib);
            list.EnableFolderCreation = true;
            list.Update();
            Context.Load(list.RootFolder);
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
        /// Copy folder to a SharePoint Online Site library
        /// </summary>
        /// <param name="folder">Folder to copy</param>
        /// <param name="list">List to copy folder to</param>
        /// <param name="localPath">Local Path selected by user</param>
        /// <returns>CopyStatus - the result from processing</returns>
        public CopyStatus copyFolderToSPO(DirectoryInfo folder, List list, string localPath)
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
            if (checkFolderExist(folderUrls.ServerRelativeUrl) == false)
            {
                //We create the folder ListITemCreationInformation
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderUrls.ItemNormalizedPath;

                //We create the folder
                ListItem newItem = list.AddItem(itemCreateInfo);

                //We update the folder metadata
                newItem["Title"] = folderUrls.ItemNormalizedPath;
                newItem["Created"] = folder.CreationTimeUtc;
                newItem["Modified"] = folder.CreationTimeUtc;
                newItem.Update();

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
        /// From MSDocs - to investigate
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="libraryName"></param>
        /// <param name="fileName"></param>
        /// <param name="fileChunkSizeInMB"></param>
        /// <returns></returns>
        public Microsoft.SharePoint.Client.File UploadFileSlicePerSlice(ClientContext ctx, string libraryName, string fileName, string itemNormalizedPath, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the name of the file.
            string uniqueFileName = Path.GetFileName(fileName);

            // Get the folder to upload into. 
            List docs = ctx.Web.Lists.GetByTitle(libraryName);
            ctx.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // File object.
            Microsoft.SharePoint.Client.File uploadFile;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            // Get the information about the folder that will hold the file.
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();


            // Get the size of the file.
            long fileSize = new FileInfo(fileName).Length;

            if (fileSize <= blockSize)
            {
                // Use regular approach.
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = itemNormalizedPath;
                    fileInfo.Overwrite = true;
                    uploadFile = docs.RootFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // Return the file object for the uploaded file.
                    return uploadFile;
                }
            }
            else
            {
                // Use large file upload approach.
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

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = itemNormalizedPath;
                                    fileInfo.Overwrite = true;
                                    uploadFile = docs.RootFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(itemNormalizedPath);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();

                                        // Return the file object for the uploaded file.
                                        return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
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

            return null;
        }


        /// <summary>
        /// Copy File to a SharePoint Online library
        /// </summary>
        /// <param name="file">File to copy</param>
        /// <param name="list">The Targeted SharePoint Online Library</param>
        /// <param name="localPath">Local Path selected by user</param>
        /// <returns>CopyStatus - The result of processing</returns>
        public CopyStatus copyFileToSPO(FileInfo file, List list, string localPath)
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
                targetFileStat = checkItemExist(itemUrl);
            }

            //If the target item does not exist, we copy the file and set metadata
            if (targetFileStat.FileFound == OnlineFileStatus.FileStatus.NotFound)
            {
                //We copy the file and set metadata
                //Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, fileUrls.ServerRelativeUrl, fileStream, false);
                UploadFileSlicePerSlice(Context, "Documents", file.FullName, fileUrls.ServerRelativeUrl);
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
                    UploadFileSlicePerSlice(Context, "Documents", file.FullName, fileUrls.ServerRelativeUrl);
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
                    UploadFileSlicePerSlice(Context, "Documents", file.FullName, fileUrls.ServerRelativeUrl);
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
        /// Verify if the item allready exist in the SharePoint Online library and retrieve the fileHash column value
        /// </summary>
        /// <param name="itemPath"></param>
        /// <returns>OnlineFileStatus (FileFound?,HashFound?,HashValue)</returns>
        private OnlineFileStatus checkItemExist(string itemPath)
        {
            //We instanciate the OnlineFileStatus object
            OnlineFileStatus status = new OnlineFileStatus();
            try
            {
                //We try to retrieve the ListItem
                ListItem file = Context.Web.GetListItem(itemPath);
                Context.Load(file);
                Context.ExecuteQuery();
                //At this point we found the file so we update the OnlineFileStatus accordingly
                status.FileFound = OnlineFileStatus.FileStatus.Found;

                //We try to retrieve the value from the hashColumn
                string fileHash = file[this.hashColumn].ToString();
                //At this point we found the file so we update the OnlineFileStatus accordingly
                status.HashFound = OnlineFileStatus.HashStatus.Found;
                status.Hash = fileHash;

                return status;
            }
            catch (ServerException ex)
            {
                //We isolate the FileNotFound exception
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    //We update the OnlineFileStatus accordingly
                    status.FileFound = OnlineFileStatus.FileStatus.NotFound;
                    status.HashFound = OnlineFileStatus.HashStatus.NotFound;
                    status.Hash = null;

                    return status;
                }
                throw ex;
            }
            catch (Exception ex)
            {
                //We isolate the Field is null exception => /!\ To improve /!\
                if (ex.HResult == -2147467261)
                {
                    //We update the OnlineFileStatus accordingly
                    status.HashFound = OnlineFileStatus.HashStatus.NotFound;
                    status.Hash = null;
                    return status;
                }
                throw ex;
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
