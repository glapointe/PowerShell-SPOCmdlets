using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class SPOWebPipeBind : CmdletPipeBind<string>
    {
        private string _surl = null;
        private Web _web = null;

        public SPOWebPipeBind(Guid inputObject)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            _web = ctx.Site.OpenWebById(inputObject);
            ctx.Load(_web, w => w.ServerRelativeUrl, w => w.HasUniqueRoleAssignments);
            ctx.ExecuteQuery();
            _surl = _web.ServerRelativeUrl;
        }
        public SPOWebPipeBind(string inputObject) : base(inputObject)
        {
            _surl = inputObject;
        }
        public SPOWebPipeBind(Uri inputObject)
        {
            _surl = inputObject.PathAndQuery;
        }
        public SPOWebPipeBind(SPOWeb inputObject)
        {
            _web = inputObject.Web;
            _surl = inputObject.ServerRelativeUrl;
        }
        public SPOWebPipeBind(Web inputObject)
        {
            _web = inputObject;
            if (!inputObject.IsPropertyAvailable("ServerRelativeUrl"))
            {
                inputObject.Context.Load(inputObject, w => w.ServerRelativeUrl);
                inputObject.Context.ExecuteQuery();
            }
            _surl = inputObject.ServerRelativeUrl;
        }

        protected override void Discover(string instance)
        {
            _surl = instance;
        }

        public override string Read()
        {
            return _surl;
        }

        public override string ToString()
        {
            return _surl;
        }
        internal void Refresh()
        {
            if (_web != null)
                new SPOWeb(_web).Refresh();
        }
    }
}
