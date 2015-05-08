using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell
{
    public class SPOSiteContext
    {
        private static SPOSiteContext _currentSiteContext = null;
        private SPOSite _site = null;
        private SPOSite _detailedSite = null;

        public SPOSiteContext(CmdletContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }
            this.Context = context;
        }

        public CmdletContext Context { get; protected set; }

        internal static SPOSiteContext CurrentSiteContext { 
            get { return _currentSiteContext; }
            set { _currentSiteContext = value; }
        }

        public string Url
        {
            get
            {
                return this.Context.Url;
            }
        }

        public SPOSite GetSite(bool detailed)
        {
            if (_detailedSite != null) return _detailedSite;
            if (_site != null && !detailed) return _site;

            var site = Context.Site;
            SPOSite.LoadSite(Context, site, detailed);

            return new SPOSite(site);
        }

    }

}
