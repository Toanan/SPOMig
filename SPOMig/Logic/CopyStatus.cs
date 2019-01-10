﻿using System;
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
        public string Type { get; set; }
        public string Path { get; set; }
        public string Status { get; set; }
        public string Comment { get; set; }
    }
}
