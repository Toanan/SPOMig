using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SPOMig
{
    /// <summary>
    /// This class is used to create result and log file
    /// </summary>
    class Reporting
    {
        #region Props
        public string LogFilePath { get; set; }
        public string ResultFilePath { get; set; }
        private enum reportFileType { Result, Log }
        #endregion

        #region Ctor
        public Reporting(string libName)
        {
            createFiles(libName);
        }
        #endregion

        #region Methods

        /// <summary>
        /// Create the result csv file and set header
        /// TODO : Create the log file
        /// </summary>
        /// <param name="libName">SharePoint Online library name</param>
        private void createFiles (string libName)
        {
            //We create the result file name and ensure the Result folder exists
            string resultFilePath = setFilePath(libName, reportFileType.Result);
            string logFilePath = setFilePath(libName, reportFileType.Log);

            //We set to FilePath properties of the Reporting object
            this.ResultFilePath = resultFilePath;
            this.LogFilePath = logFilePath;

            //We create the result file and set the header
            var resultHeader = new StringBuilder();
            var header = "Name,Type,Path,Status,Comment";
            resultHeader.AppendLine(header);
            File.WriteAllText(resultFilePath, resultHeader.ToString(), Encoding.UTF8);

            //We create the log file TODO

        }

        /// <summary>
        /// Write the processing result on the result csv file
        /// </summary>
        /// <param name="copyStatus">CopyStatus object</param>
        public void writeResult (CopyStatus copyStatus)
        {
            var csv = new StringBuilder();
            string result = $"{copyStatus.Name},{copyStatus.Type},{copyStatus.Path},{copyStatus.Status},{copyStatus.Comment}";
            csv.AppendLine(result);
            File.AppendAllText(this.ResultFilePath, csv.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// TODO
        /// </summary>
        /// <param name="log"></param>
        public void writeLog (string log)
        {
            var logLine = new StringBuilder();
            logLine.AppendLine(log);
            File.AppendAllText(this.LogFilePath, logLine.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Return the repport file path and ensure it exists
        /// </summary>
        /// <param name="libName"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private string setFilePath(string libName, reportFileType type)
        {
            //We create name of the file
            DateTime now = DateTime.Now;
            var date = now.ToString("yyyy-MM-dd-HH-mm-ss");
            string FileName = $"{libName}-{date}";
            //string appPath = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appName = "SPOMig";

            //We loop the file type to return the path accordingly
            switch (type)
            {
                case reportFileType.Result:

                    string resultFilePath = $"{appPath}/{appName}/Results/{FileName}.csv";
                    //We ensure path exists
                    if (!Directory.Exists($"{appPath}/{appName}/Results/")) Directory.CreateDirectory($"{appPath}/{appName}/Results/");
                    return resultFilePath;

                case reportFileType.Log:

                    string logFilePath = $"{appPath}/{appName}/Logs/{FileName}.log";
                    //We ensure path exists
                    if (!Directory.Exists($"{appPath}/{appName}/Logs/")) Directory.CreateDirectory($"{appPath}/{appName}/Logs/");
                    return logFilePath;

                default:

                    throw new NotImplementedException();
            }
        }
        
        #endregion
    }
}
