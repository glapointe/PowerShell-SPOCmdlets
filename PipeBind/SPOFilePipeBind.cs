using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOFilePipeBind
    {
        private string _serverRelativeUrl;

        public SPOFilePipeBind(Microsoft.SharePoint.Client.File file)
        {
            if (!file.IsPropertyAvailable("ServerRelativeUrl"))
            {
                file.Context.Load(file, f => f.ServerRelativeUrl);
                file.Context.ExecuteQuery();
            }
            _serverRelativeUrl = file.ServerRelativeUrl;
        }

        public SPOFilePipeBind(SPOFile file)
        {
            _serverRelativeUrl = file.ServerRelativeUrl;
        }

        public SPOFilePipeBind(SPOWikiPage wikiPage) : this((SPOFile)wikiPage) { }

        public SPOFilePipeBind(string serverRelativeUrl)
        {
            _serverRelativeUrl = serverRelativeUrl;
        }

        public SPOFile Read(Web web)
        {
            SPOWeb spoWeb = new SPOWeb(web);
            return Read(spoWeb);
        }

        public SPOFile Read(SPOWeb spoWeb)
        {
            if (string.IsNullOrEmpty(_serverRelativeUrl))
                throw new ArgumentNullException("The server relative URL was not provided.");

            return spoWeb.GetFileByServerRelativeUrl(_serverRelativeUrl);
        }
    }
}
