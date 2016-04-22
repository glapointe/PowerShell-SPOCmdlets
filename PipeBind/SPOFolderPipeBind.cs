using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOFolderPipeBind
    {
        private string _serverRelativeUrl;

        public SPOFolderPipeBind(Microsoft.SharePoint.Client.Folder folder)
        {
            if (!folder.IsPropertyAvailable("ServerRelativeUrl"))
            {
                folder.Context.Load(folder, f => f.ServerRelativeUrl);
                folder.Context.ExecuteQuery();
            }
            _serverRelativeUrl = folder.ServerRelativeUrl;
        }

        public SPOFolderPipeBind(SPOFolder folder)
        {
            _serverRelativeUrl = folder.ServerRelativeUrl;
        }

        public SPOFolderPipeBind(string serverRelativeUrl)
        {
            _serverRelativeUrl = serverRelativeUrl;
        }

        public SPOFolder Read(Web web)
        {
            SPOWeb spoWeb = new SPOWeb(web);
            return Read(spoWeb);
        }

        public SPOFolder Read(SPOWeb spoWeb)
        {
            if (string.IsNullOrEmpty(_serverRelativeUrl))
                throw new ArgumentNullException("The server relative URL was not provided.");

            return spoWeb.GetFolderByServerRelativeUrl(_serverRelativeUrl);
        }
    }
}
