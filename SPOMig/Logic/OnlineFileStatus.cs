using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    /// <summary>
    /// This Class is used to store feedback when processing file comparison in the file copy process
    /// </summary>
    class OnlineFileStatus
    {
        public FileStatus FileFound { get; set; }
        public HashStatus HashFound { get; set; }
        public string Hash { get; set; }
        public enum FileStatus { Found, NotFound }
        public enum HashStatus { Found, NotFound }
    }
}
