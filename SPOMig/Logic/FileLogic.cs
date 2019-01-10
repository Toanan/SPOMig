using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace SPOMig
{
    class FileLogic
    {
        #region Props
        public string LocalPath { get; set; }
        #endregion

        #region Ctor
        public FileLogic(string path)
        {
            this.LocalPath = path;
        }
        #endregion

        #region Methods

        /// <summary>
        /// Retrieve all file information recursively from a local path
        /// </summary>
        /// <returns>List<FileInfo></returns>
        public List<FileInfo> getFiles()
        {
            //We retrieve the sub dirinfos
            List<DirectoryInfo> sourceFolders = getSourceFolders();

            //We create the files fileinfo object
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
        /// Retrieve folders from local directory
        /// </summary>
        /// <param name="url"></param>
        public List<DirectoryInfo> getSourceFolders()
        {
            //We retrieve all directories from the local path in an array
            string[] Folders = Directory.GetDirectories(LocalPath, "*.*", SearchOption.AllDirectories);
            
            //We create the list of all directoryinfo
            List<DirectoryInfo> folders = new List<DirectoryInfo>();
            //We create the source rootFolder DirInfo and add it to the top of the list
            DirectoryInfo rootFolder = new DirectoryInfo(LocalPath);
            folders.Add(rootFolder);

            //We loop to populate directoryinfo list from directory path
            foreach (string folder in Folders)
            {
                DirectoryInfo di = new DirectoryInfo(folder);
                folders.Add(di);
            }

            return folders;
        }

        /// <summary>
        /// Retrieve local files in a folder
        /// </summary>
        /// <param name="url"></param>
        private List<FileInfo> getLocalFileInFolder(string path)
        {
            //We retrieve file path from the directory path
            string[] Files = Directory.GetFiles(path, "*.*", SearchOption.TopDirectoryOnly);
            //We create the list to store files info
            List<FileInfo> files = new List<FileInfo>();

            //We loop to populate fileinfo from file path
            foreach (string File in Files)
            {
                FileInfo fi = new FileInfo(File);
                files.Add(fi);
            }

            return files;
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

                return MakeHashString(hasher.Hash);

            }

        }

        #endregion
    }
}
