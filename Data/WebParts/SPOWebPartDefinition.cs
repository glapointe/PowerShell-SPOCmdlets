using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.WebParts;

namespace Lapointe.SharePointOnline.PowerShell.Data.WebParts
{
    public class SPOWebPartDefinition : SPOClientObject
    {
        private WebPartDefinition _webPartDefinition = null;

        public SPOWebPartDefinition(WebPartDefinition webPartDefinition)
        {
            _webPartDefinition = webPartDefinition;

        }

        public WebPartDefinition WebPartDefinition { get { return _webPartDefinition; } }

        public Guid Id
        {
            get
            {
                if (_webPartDefinition.IsPropertyAvailable("Id"))
                    return _webPartDefinition.Id;
                return Guid.Empty;
            }
        }
        

        public void CloseWebPart()
        {
            _webPartDefinition.CloseWebPart();
            _webPartDefinition.SaveWebPartChanges();
            _webPartDefinition.Context.ExecuteQuery();
        }

        public void DeleteWebPart()
        {
            _webPartDefinition.DeleteWebPart();
            _webPartDefinition.Context.ExecuteQuery();
        }

        public void MoveWebPartTo(string zoneID, int zoneIndex)
        {
            _webPartDefinition.MoveWebPartTo(zoneID, zoneIndex);
            _webPartDefinition.SaveWebPartChanges();
            _webPartDefinition.Context.ExecuteQuery();
        }

        public void OpenWebPart()
        {
            _webPartDefinition.OpenWebPart();
            _webPartDefinition.SaveWebPartChanges();
            _webPartDefinition.Context.ExecuteQuery();
        }

        internal void SaveWebPartChanges()
        {
            _webPartDefinition.SaveWebPartChanges();
            _webPartDefinition.Context.ExecuteQuery();
        }

        private SPOWebPart _webPart = null;
        public SPOWebPart WebPart
        {
            get
            {
                if (_webPart != null) return _webPart;

                var webPart = _webPartDefinition.WebPart;
                SPOWebPart.LoadWebPart(SPOSiteContext.CurrentSiteContext.Context, webPart);
                _webPart = new SPOWebPart(webPart, this);
                return _webPart;
            }
        }
    }
}
