using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    /// <summary>
    /// This class is used to provide feedback to the writeResult method of the report class
    /// </summary>
    class CopyStatus
    {
        public string Name { get; set; }
        public ItemType Type { get; set; }
        public string Path { get; set; }
        public ItemStatus Status { get; set; }
        public string Comment { get; set; }
        public enum ItemType { File, Folder }
        public enum ItemStatus { Created, Skiped, Uploaded, Overwrited, Error, Deleted, InProgress }
    }
}
