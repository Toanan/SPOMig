using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

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
        /// Retrive all file information recursively from a path
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
                List<FileInfo> Currentfiles = getSourceFiles(directory.FullName);
                foreach (FileInfo fi in Currentfiles)
                {
                    files.Add(fi);
                }
            }
            return files;
        }

        /// <summary>
        /// Retrive items from local directory
        /// </summary>
        /// <param name="url"></param>
        public List<DirectoryInfo> getSourceFolders()
        {
            // TODO ADD the root directory !!
            string[] Folders = Directory.GetDirectories(LocalPath, "*.*", SearchOption.AllDirectories);
            //We create the list to put all directories
            List<DirectoryInfo> folders = new List<DirectoryInfo>();
            //We create the source rootFolder DirInfo and add it to the top of the list
            DirectoryInfo rootFolder = new DirectoryInfo(LocalPath);
            folders.Add(rootFolder);

            //We loop to populate directory info from directory path
            foreach (string folder in Folders)
            {
                DirectoryInfo di = new DirectoryInfo(folder);
                folders.Add(di);
            }

            return folders;
        }

        /// <summary>
        /// Retrive items from local directory
        /// </summary>
        /// <param name="url"></param>
        private List<FileInfo> getSourceFiles(string path)
        {
            //We retrive file path from the directory path
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

        #endregion
    }
}
