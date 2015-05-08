using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client.WebParts;

namespace Lapointe.SharePointOnline.PowerShell.Data.WebParts
{
    public class SPOWebPart : SPOClientObject
    {
        private WebPart _webPart = null;
        private SPOWebPartDefinition _webPartDefinition = null;
        private bool _hasChanges = false;

        public SPOWebPart(WebPart webPart, SPOWebPartDefinition webPartDefinition)
        {
            _webPart = webPart;
            _webPartDefinition = webPartDefinition;
        }

        public WebPart WebPart { get { return _webPart; } }

        public void Update()
        {
            if (_hasChanges)
            {
                _webPartDefinition.SaveWebPartChanges();
                _hasChanges = false;
            }
        }

        public bool Hidden
        {
            get
            {
                if (_webPart.IsPropertyAvailable("Hidden"))
                    return _webPart.Hidden;
                return false;
            }
            set
            {
                _webPart.Hidden = value;
                _hasChanges = true;
            }
        }

        public bool IsClosed
        {
            get
            {
                if (_webPart.IsPropertyAvailable("IsClosed"))
                    return _webPart.IsClosed;
                return false;
            }
        }

        private SPOPropertyBag _properties = null;
        public SPOPropertyBag Properties
        {
            get
            {
                if (_properties == null && _webPart.IsObjectPropertyInstantiated("Properties"))
                    _properties = new SPOPropertyBag(this, _webPart.Properties);
                return _properties;
            }
        }
        
        public string Subtitle
        {
            get
            {
                if (_webPart.IsPropertyAvailable("Subtitle"))
                    return _webPart.Subtitle;
                return null;
            }
        }

        public string Title
        {
            get
            {
                if (_webPart.IsPropertyAvailable("Title"))
                    return _webPart.Title;
                return null;
            }
            set
            {
                _webPart.Title = value;
                _hasChanges = true;
            }
        }

        public string TitleUrl
        {
            get
            {
                if (_webPart.IsPropertyAvailable("TitleUrl"))
                    return _webPart.TitleUrl;
                return null;
            }
            set
            {
                _webPart.TitleUrl = value;
                _hasChanges = true;
            }
        }

        public int ZoneIndex
        {
            get
            {
                if (_webPart.IsPropertyAvailable("ZoneIndex"))
                    return _webPart.ZoneIndex;
                return 0;
            }
        }



        internal static void LoadWebPart(CmdletContext ctx, WebPart webPart)
        {
            ctx.Load(webPart,
                wp => wp.Hidden,
                wp => wp.IsClosed,
                wp => wp.Properties,
                wp => wp.Subtitle,
                wp => wp.Title,
                wp => wp.TitleUrl,
                wp => wp.ZoneIndex);
            ctx.ExecuteQuery();
        }
    }
}
