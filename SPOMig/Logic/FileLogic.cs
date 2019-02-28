using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Cryptography;
using System;
using System.Xml;
using System.Xml.Linq;

namespace SPOMig
{
    /// <summary>
    /// This class is used to interact with the local FileSystem
    /// It provides methods to retrive a list of DirectoryInfo and FileInfo from a path
    /// It also provide a method to compute file hash
    /// </summary>
    static class FileLogic
    {

        #region Methods

        /// <summary>
        /// Retrieve all file information recursively from a local path
        /// </summary>
        /// <returns>List<FileInfo></returns>
        public static List<FileInfo> getFiles(string localPath)
        {
            //We retrieve the sub dirinfos
            List<DirectoryInfo> sourceFolders = getSourceFolders(localPath);

            //We create the list of fileinfo to retrieve
            List<FileInfo> files = new List<FileInfo>();

            //And loop inside all dir to retrieve the files fileinfo
            foreach (DirectoryInfo directory in sourceFolders)
            {
                List<FileInfo> Currentfiles = getLocalFileInFolder(directory.FullName);
                foreach (FileInfo fi in Currentfiles)
                {
                    files.Add(fi);
                }
            }
            return files;
        }

        /// <summary>
        /// Retrieve all DirectoryInfo from a local path
        /// </summary>
        /// <param name="localPath"></param>
        public static List<DirectoryInfo> getSourceFolders(string localPath)
        {
            //We retrieve all directories path from the local path in an array
            string[] foldersPath = Directory.GetDirectories(localPath, "*.*", SearchOption.AllDirectories);
            
            //We create the list of all Directoryinfo to retrieve
            List<DirectoryInfo> folders = new List<DirectoryInfo>();

            //We add the rootFolder from the local path
            DirectoryInfo rootFolder = new DirectoryInfo(localPath);
            folders.Add(rootFolder);

            //We loop the foldersPath array to retrive all the DirectoryInfo
            foreach (string folder in foldersPath)
            {
                DirectoryInfo di = new DirectoryInfo(folder);
                folders.Add(di);
            }

            return folders;
        }

        /// <summary>
        /// Retrieve FileInfo from a folder
        /// </summary>
        /// <param name="folderPath"></param>
        private static List<FileInfo> getLocalFileInFolder(string folderPath)
        {
            //We retrieve all file path from the local directory path
            string[] filesPath = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly);
            
            //We create the list of all FilesInfo to retrieve
            List<FileInfo> files = new List<FileInfo>();

            //We loop the filePath array to retrieve all the FileInfo
            foreach (string File in filesPath)
            {
                FileInfo fi = new FileInfo(File);
                files.Add(fi);
            }

            return files;
        }

        /// <summary>
        /// Convert a hashBytes to string 
        /// </summary>
        /// <param name="hashBytes">the array of byte to convert</param>
        /// <returns>string</returns>
        private static string convertHashToString(byte[] hashBytes)
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
        public static string hashFromLocal(FileStream localFileStream)
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
                return convertHashToString(hasher.Hash);
            }
        }

        /// <summary>
        /// Ensure or create the xml config file with default settings (located in %appdata%)
        /// </summary>
        public static void ensureConfigFileExists()
        {
            string xmlPath = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "\\SPOMig\\cfg.xml");

            if (!File.Exists(xmlPath))
            {
                using (XmlWriter writer = XmlWriter.Create(xmlPath))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Settings");

                    writer.WriteAttributeString("HashColumn", "FileHash");
                    writer.WriteAttributeString("appID", "");
                    writer.WriteAttributeString("appSecret", "");
                    writer.WriteAttributeString("Threads", "1");

                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
        }

        /// <summary>
        /// retrieve an xml attribute value from the cfg.xml setting file
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string getXMLSettings(string key)
        {
            string xmlPath = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "\\SPOMig\\cfg.xml");
            string value = "";

            using (XmlReader xmlReader = XmlReader.Create(xmlPath))
            {
                while (xmlReader.Read())
                {
                    if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "Settings"))
                    {
                        if (xmlReader.HasAttributes)

                            value = xmlReader.GetAttribute(key);
                    }
                }
            }
            return value;
        }

        /// <summary>
        /// Set an xml attribute value from the cfg.xml setting file
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public static void setXMLSettingsAttribute(string key, string value)
        {
            string xmlFilePath = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "\\SPOMig\\cfg.xml");

            string newValue = string.Empty;
            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.Load(xmlFilePath);

            foreach (XmlNode node in xmlDoc.ChildNodes)
            {
                if (node.Name == "Settings")
                {
                    node.Attributes[key].Value = value;
                }
            }
            xmlDoc.Save(xmlFilePath);
        }
        #endregion
    }
}
