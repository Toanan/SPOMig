using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SPOMig
{

    class WriteRepport
    {

        #region Props
        public string LogFileName { get; set; }
        public string LogFilePath { get; set; }
        public string ResultFileName { get; set; }
        public string ResultFilePath { get; set; }
        #endregion

        #region Ctor
        public WriteRepport(string libName)
        {
            createResultFile(libName);
        }
        #endregion

        #region Methods
        private void createResultFile (string libName)
        {
            //We create the path of the csv file
            DateTime now = DateTime.Now;
            var date = now.ToString("yyyy-MM-dd-HH-mm-ss");
            string csvFileName = $"{libName}-{date}";
            var appPath = AppDomain.CurrentDomain.BaseDirectory;
            var csvfilePath = $"{appPath}/Results/{csvFileName}{date}.csv";
            if (!Directory.Exists($"{appPath}/Results/")) Directory.CreateDirectory($"{appPath}/Results/");

            this.ResultFilePath = csvfilePath;

            //We create the result file by writing the header
            var csv = new StringBuilder();
            var header = "Name,Type,OnlinePath,Status,Comment";
            csv.AppendLine(header);
            File.WriteAllText(csvfilePath, csv.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Write result when copying file and folder
        /// </summary>
        /// <param name="result"></param>
        public void writeResult (string result)
        {
            var csv = new StringBuilder();
            csv.AppendLine(result);
            File.AppendAllText(this.ResultFilePath, csv.ToString(), Encoding.UTF8);
        }

        private void createLogFile (string libName)
        {
            //We create the path and Name of the log file
            DateTime now = DateTime.Now;
            var date = now.ToString("yyyy-MM-dd-HH-mm-ss");
            string logFileName = $"{libName}-{date}";
            var appPath = AppDomain.CurrentDomain.BaseDirectory;
            var logfilePath = $"{appPath}/Logs/{logFileName}{date}.csv";
            if (!Directory.Exists($"{appPath}/Logs/")) Directory.CreateDirectory($"{appPath}/Logs/");

            this.LogFilePath = logfilePath;
        }

        public void writeLog (string log)
        {
            var logLine = new StringBuilder();
            logLine.AppendLine(log);
            File.AppendAllText(this.LogFilePath, logLine.ToString(), Encoding.UTF8);
        }
        #endregion



    }
}
