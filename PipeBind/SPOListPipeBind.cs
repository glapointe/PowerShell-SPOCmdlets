using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOListPipeBind
    {
        private string _listUrl;
        private Guid _listId;
        private string _listTitle;
        private List _list = null;

        public SPOListPipeBind(SPOList spoList)
        {
            if ((spoList == null) || spoList.Id == Guid.Empty)
            {
                throw new ArgumentNullException("The List ID must be specified.");
            }
            _listId = spoList.Id;
            _list = spoList.List;
        }
        public SPOListPipeBind(List list)
        {
            if ((list == null) || list.Id == Guid.Empty)
            {
                throw new ArgumentNullException("The List ID must be specified.");
            }
            _listId = list.Id;
            _list = list;
        }
        public SPOListPipeBind(Guid listId)
        {
            _listId = listId;
        }

        public SPOListPipeBind(string urlOrTitle)
        {
            try
            {
                if (urlOrTitle.StartsWith("/"))
                    _listUrl = urlOrTitle;
                else
                {
                    SPOSiteHelper.ValidateUrl("SPOListPipeBind Url", urlOrTitle);
                    _listUrl = urlOrTitle.Trim();
                }
            }
            catch (ArgumentException)
            {
                _listTitle = urlOrTitle;
            }
        }

        public SPOListPipeBind(Uri uri)
        {
            SPOSiteHelper.ValidateUri("SPOListPipeBind Uri", uri);
            _listUrl = uri.AbsoluteUri;
        }

        public string Url
        {
            get { return _listUrl; }
        }

        public string ListTitle
        {
            get { return _listTitle; }
        }

        public Guid ListId
        {
            get { return _listId; }
        }

        public SPOList Read(Web web, bool detailed)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            List list = null;
            if (this.ListId != Guid.Empty)
            {
                list = web.Lists.GetById(this.ListId);
            }
            else if (!string.IsNullOrEmpty(this.ListTitle))
            {
                list = web.Lists.GetByTitle(this.ListTitle);
            }
            else if (!string.IsNullOrEmpty(this.Url))
            {
                var lists = web.Lists;
                ctx.Load(lists, l => l.Include(a => a.Title, a => a.Id, a => a.RootFolder));
                ctx.ExecuteQuery();

                foreach (var list2 in lists)
                {
                    string pathAndQuery = HttpUtility.UrlPathEncode(this.Url, true);
                    if (!this.Url.StartsWith("/"))
                        pathAndQuery = new Uri(this.Url).PathAndQuery;

                    if (HttpUtility.UrlPathEncode(list2.RootFolder.ServerRelativeUrl, true).ToLower().TrimEnd('/') == pathAndQuery.ToLower().TrimEnd('/'))
                    {
                        list = list2;
                        break;
                    }
                }
            }
            if (list != null)
            {
                _list = list;
                SPOList.LoadList(ctx, list, detailed);

                return new SPOList(list);
            }
            return null;
        }
        public override string ToString()
        {
            if (this.ListId != Guid.Empty)
                return this.ListId.ToString();
            else if (!string.IsNullOrEmpty(this.ListTitle))
                return this.ListTitle;
            else if (!string.IsNullOrEmpty(this.Url))
                return this.Url;
            return "List";
        }
        public void Refresh()
        {
            if (_list != null)
                new SPOList(_list).Refresh();
        }
    }
}
