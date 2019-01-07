using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Security.Cryptography;

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
            using (FileStream fileStream = new FileStream(file.FullName, FileMode.Open))
            {
                string libURL = list.RootFolder.ServerRelativeUrl.ToString();
                string localPathNormalized = localPath.Replace("/", "\\");
                string filePath = file.FullName.Replace("/", "\\");
                string fileNormalizedPath = filePath.Replace(localPathNormalized, "");
                string fileNormalizedPathfinal = fileNormalizedPath.Replace("\\", "/");
                string serverRelativeURL = libURL + "/" + fileNormalizedPathfinal;

                Web site = Context.Web;
                Context.Load(site, s => s.Url);
                Context.ExecuteQuery();
                string itemUrl = site.Url + "/" + list.RootFolder.Name + "/" + fileNormalizedPathfinal;

                long targetLenght = checkItemExist(itemUrl);

                if (targetLenght == 0)
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, false);
                    string currentHash = hashFromLocal(fileStream);

                    Microsoft.SharePoint.Client.ListItem currentOnlinefile = Context.Web.GetListItem(serverRelativeURL);
                    currentOnlinefile["test"] = currentHash;
                    currentOnlinefile.Update();
                    Context.ExecuteQuery();

                    return true;
                }
                else //File allready exist => compare hash
                {
                    long localLenght = file.Length;

                    if (localLenght == targetLenght)
                    {
                        return false;
                    }
                    else
                    {
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(Context, serverRelativeURL, fileStream, true);
                        return true;
                    }
                }
            }
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

        /// <summary>
        /// Compute a hash string from hashBytes
        /// </summary>
        /// <param name="hashBytes"></param>
        /// <returns>hash string</returns>
        private static string MakeHashString(byte[] hashBytes)
        {
            StringBuilder hash = new StringBuilder(32);

            foreach (byte b in hashBytes)
                hash.Append(b.ToString("X2").ToLower());

            return hash.ToString();
        }

        /// <summary>
        /// Compute the hash string from a filestream
        /// </summary>
        /// <param name="localFileStream"></param>
        /// <returns>hash string</returns>
        private string hashFromLocal (FileStream localFileStream)
        {
            byte[] buffer;
            int byteRead;
            long size;
            long totalByteRead = 0;

            Stream file = localFileStream;
            

            size = file.Length;

            using (HashAlgorithm hasher = MD5.Create())
            {
                do
                {
                    buffer = new byte[4096];

                    byteRead = file.Read(buffer, 0, buffer.Length);

                    totalByteRead += byteRead;

                    hasher.TransformBlock(buffer, 0, byteRead, null, 0);

                }
                while (byteRead != 0);

                hasher.TransformFinalBlock(buffer, 0, 0);

                return MakeHashString(hasher.Hash);

            }
            
        }
        #endregion

    }
}
