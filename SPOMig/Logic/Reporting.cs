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
        public Reporting(string libName, string siteUrl)
        {
            createFiles(libName, siteUrl);
        }
        #endregion

        #region Methods

        /// <summary>
        /// Create the result csv file and set header
        /// </summary>
        /// <param name="libName">SharePoint Online library name</param>
        private void createFiles (string libName, string siteUrl)
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

            //We create the log file by writing process start
            var logStartMessage = $"[Process beggin]Destination Site : {siteUrl} | Destination Library : {libName}";
            CopyLog log = new CopyLog(logStartMessage);
            writeLog(log);
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
        /// Write logs on the log file
        /// </summary>
        /// <param name="log"></param>
        public void writeLog (CopyLog log)
        {
            //We retrieve the date 
            DateTime date = DateTime.Now;
            string formatedDate = date.ToString("yyyy-MM-dd-HH-mm-ss");

            //We iterate the ActionStatus value to render the log
            switch (log.ActionStatus)
            {
                case CopyLog.Status.Verbose:
                    var verboseLogLine = new StringBuilder();
                    verboseLogLine.AppendLine($"{formatedDate}-[{log.Action} - {log.ActionStatus}] #Path : {log.ItemPath}");
                    File.AppendAllText(this.LogFilePath, verboseLogLine.ToString(), Encoding.UTF8);
                    break;

                case CopyLog.Status.Empty:
                    var emptyLogLine = new StringBuilder();
                    emptyLogLine.AppendLine("------------------------------------------------");
                    emptyLogLine.AppendLine($"{formatedDate}-{log.Comment}");
                    emptyLogLine.AppendLine("------------------------------------------------");
                    File.AppendAllText(this.LogFilePath, emptyLogLine.ToString(), Encoding.UTF8);
                    break;

                default:
                    //Do we display a comment ?
                    if (string.IsNullOrWhiteSpace(log.Comment))
                    {
                        var defaultLogLine = new StringBuilder();
                        defaultLogLine.AppendLine($"{formatedDate}-[{log.Action} - {log.ActionStatus}] #Path : {log.ItemPath}");
                        File.AppendAllText(this.LogFilePath, defaultLogLine.ToString(), Encoding.UTF8);
                    }
                    else
                    {
                        var defaultLogLine = new StringBuilder();
                        defaultLogLine.AppendLine($"{formatedDate}-[{log.Action} - {log.ActionStatus}] #Path : {log.ItemPath} #Message : {log.Comment}");
                        File.AppendAllText(this.LogFilePath, defaultLogLine.ToString(), Encoding.UTF8);
                    }
                    
                    break;
            }
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
