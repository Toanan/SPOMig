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
        public enum Status { OK, Warning, Error}
        public string ItemPath { get; set; }
        public string Action { get; set; }
        public string Path { get; set; }
        public string Comment { get; set; }
    }
}
