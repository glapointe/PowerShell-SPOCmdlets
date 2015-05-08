using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Principals;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class SPOGroupPipeBind
    {
        private SPOGroup _group = null;

        public SPOGroupPipeBind(Group group)
        {
            _group = new SPOGroup(group);
        }

        public SPOGroupPipeBind(SPOGroup group)
        {
            _group = group;
        }

        public SPOGroupPipeBind(int id)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var group = ctx.Site.RootWeb.SiteGroups.GetById(id);
            ctx.Load(group);
            ctx.ExecuteQuery();
            _group = new SPOGroup(group);
        }

        public SPOGroupPipeBind(string name)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var group = ctx.Site.RootWeb.SiteGroups.GetByName(name);
            ctx.Load(group);
            ctx.ExecuteQuery();
            _group = new SPOGroup(group);
        }

        public SPOGroup Read()
        {
            return _group;
        }
    }
}
