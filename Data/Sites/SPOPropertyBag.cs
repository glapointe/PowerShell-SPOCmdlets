using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.WebParts;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Sites
{
    public class SPOPropertyBag : SPOClientObject
    {
        private SPOWeb _web = null;
        private SPOFolder _folder = null;
        private SPOWebPart _webPart = null;
        private PropertyValues _pb = null;
        private bool _hasChanges = false;

        public SPOPropertyBag(SPOFolder folder, PropertyValues properties)
        {
            _folder = folder;
            _pb = properties;
        }
        public SPOPropertyBag(SPOWeb web, PropertyValues properties)
        {
            _web = web;
            _pb = properties;
        }
        public SPOPropertyBag(SPOWebPart webPart, PropertyValues properties)
        {
            _webPart = webPart;
            _pb = properties;
        }

        public Dictionary<string, object> FieldValues
        {
            get
            {
                return _pb.FieldValues;
            }
        }

        public object this[string fieldName]
        {
            get
            {
                return _pb[fieldName];
            }
            set
            {
                _pb[fieldName] = value;
                _hasChanges = true;
            }
        }

        public void SetFieldValue(string fieldName, object value)
        {
            _pb[fieldName] = value;
            _hasChanges = true;
        }

        public void Update()
        {
            if (_hasChanges)
            {
                if (_web != null)
                    _web.Web.Update();
                if (_folder != null)
                    _folder.Folder.Update();
                if (_webPart != null)
                {
                    _webPart.Update();
                    _hasChanges = false;
                    return;
                }

                _pb.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }
    }
}
