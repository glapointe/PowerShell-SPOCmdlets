using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOContentTypePipeBind
    {
        private string _name;
        private string _contentTypeId;
        private SPOContentType _spoContentType;
        private ContentType _contentType;

        public SPOContentTypePipeBind(SPOContentType spoContentType)
        {
            if ((spoContentType == null))
            {
                throw new ArgumentNullException("The Content Type must be specified.");
            }
            _spoContentType = spoContentType;
            _contentType = spoContentType.ContentType;
            _name = spoContentType.Name;
            _contentTypeId = spoContentType.Id.StringValue;
        }
        public SPOContentTypePipeBind(ContentType contentType)
        {
            if ((contentType == null))
            {
                throw new ArgumentNullException("The Content Type must be specified.");
            }
            _spoContentType = new SPOContentType(contentType);
            _contentType = contentType;
            _name = contentType.Name;
            _contentTypeId = contentType.Id.StringValue;
        }
        public SPOContentTypePipeBind(string name)
        {
            if (name.ToLower().StartsWith("0x0"))
                _contentTypeId = name;
            else
                _name = name;
        }
        public SPOContentTypePipeBind(ContentTypeId contentTypeId)
        {
            _contentTypeId = contentTypeId.StringValue;
        }

        public SPOContentType Read(Web web)
        {
            return Read(web.AvailableContentTypes);
        }
        public SPOContentType Read(List list)
        {
            return Read(list.ContentTypes);
        }

        private SPOContentType Read(ContentTypeCollection contentTypes)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            ContentType contentType = null;
            if (_contentTypeId != null)
            {
                contentType = contentTypes.GetById(_contentTypeId);
            }
            else if (!string.IsNullOrEmpty(_name))
            {
                ctx.Load(contentTypes);
                ctx.ExecuteQuery();

                foreach (ContentType ct in contentTypes)
                {
                    if (ct.Name.ToLower() == _name.ToLower())
                    {
                        contentType = ct;
                        break;
                    }
                }

            }
            if (contentType != null)
            {
                SPOContentType.LoadContentType(ctx, contentType);
                if (contentType.ServerObjectIsNull.Value)
                    return null;

                return new SPOContentType(contentType);
            }

            return null;
        }
        public void Refresh()
        {
            if (_spoContentType != null)
                _spoContentType.Refresh();
            else if (_contentType != null)
                new SPOContentType(_contentType).Refresh();
        }
    }
}
