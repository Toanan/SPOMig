using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    class FileMapping
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public DateTime Modified { get; set; }
        public DateTime Created { get; set; }
        public string Owner { get; set; }
        public Type ItemType { get; set; }
        public enum Type { File, Folder }

    }
}
