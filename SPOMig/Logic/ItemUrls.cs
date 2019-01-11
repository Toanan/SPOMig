using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOMig
{
    /// <summary>
    /// This Class is used to store the URLs necessary for processing files and folders
    /// </summary>
    class ItemURLs
    {
        public string ItemNormalizedPath { get; set; }
        public string ServerRelativeUrl { get; set; }
    }
}
