using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SPOMig
{
    class SPOLogic
    {
        public ClientContext Context { get; set; }

        public SPOLogic(ClientContext ctx)
        {
            this.Context = ctx;
        }

        public ListCollection getLists()
        {
            using (Context)
            {
                ListCollection Libraries = Context.Web.Lists;
                Context.Load(Libraries);
                Context.ExecuteQuery();
                return Libraries;
            }
        }
    }
}
