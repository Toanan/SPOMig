using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    /// <summary>
    /// This class is designed to store items to process
    /// </summary>
    class FoldersToProcess
    {
        public ItemURLs ItemUrls { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
    }
}
