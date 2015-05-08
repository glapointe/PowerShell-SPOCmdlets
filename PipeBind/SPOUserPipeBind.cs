using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Principals;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class SPOUserPipeBind
    {
        private SPOUser _user = null;

        public SPOUserPipeBind(User user)
        {
            _user = new SPOUser(user);
        }

        public SPOUserPipeBind(SPOUser user)
        {
            _user = user;
        }

        public SPOUserPipeBind(int id)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var user = ctx.Site.RootWeb.SiteUsers.GetById(id);
            ctx.Load(user);
            ctx.ExecuteQuery();
            _user = new SPOUser(user);
        }

        public SPOUserPipeBind(string email)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var user = ctx.Site.RootWeb.SiteUsers.GetByEmail(email);
            ctx.Load(user);
            ctx.ExecuteQuery();
            _user = new SPOUser(user);
        }

        public SPOUser Read()
        {
            return _user;
        }
    }
}
