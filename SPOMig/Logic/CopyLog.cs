using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    /// <summary>
    /// This class is used to provide feedback to the writeLog method of the report class
    /// </summary>
    class CopyLog
    {
        #region Props
        public enum Status { OK, Warning, Error, Verbose , Empty}
        public Status ActionStatus { get; set; }
        public string ItemPath { get; set; }
        public string Action { get; set; }
        public string Comment { get; set; }
        #endregion

        #region Ctor
        public CopyLog(Status result, string action, string itemPath, string comment)
        {
            this.ActionStatus = result;
            this.ItemPath = itemPath;
            this.Action = action;
            this.Comment = comment;
        }

        public CopyLog(string comment)
        {
            this.Comment = comment;
            this.ActionStatus = Status.Empty;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Update and write log
        /// </summary>
        /// <param name="result">Result of the processing</param>
        /// <param name="action">Action to process</param>
        /// <param name="itemPath">Path of the item to process</param>
        /// <param name="comment">Comment for the processing</param>
        public void update(Status result, string action, string itemPath, string comment)
        {
            this.ActionStatus = result;
            this.ItemPath = itemPath;
            this.Action = action;
            this.Comment = comment;
        }

        /// <summary>
        /// Update and write log
        /// Overload to show only section Title
        /// </summary>
        /// <param name="comment">Name of the section</param>
        public void update(string comment)
        {
            this.ActionStatus = CopyLog.Status.Empty;
            this.ItemPath = "";
            this.Action = "";
            this.Comment = comment;
        }
        #endregion

    }
}
