using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOListItem : SPOClientObject
    {
        private ListItem _item = null;
        private bool _hasChanges = false;
        private SPOContentType _ct = null;

        public SPOListItem(ListItem item)
        {
            _item = item;
        }

        public ListItem ListItem { get { return _item; } }

        public int Id
        {
            get
            {
                if (_item.IsPropertyAvailable("Id"))
                    return _item.Id;
                return 0;
            }
        }

        public string DisplayName
        {
            get
            {
                if (_item.IsPropertyAvailable("DisplayName"))
                    return _item.DisplayName;
                return null;
            }
        }

        public DateTime Created
        {
            get
            {
                if (_item.FieldValues.ContainsKey("Created"))
                    return (DateTime)_item.FieldValues["Created"];
                return DateTime.MinValue;
            }
        }

        public DateTime Modified
        {
            get
            {
                if (_item.FieldValues.ContainsKey("Modified"))
                    return (DateTime)_item.FieldValues["Modified"];
                return DateTime.MinValue;
            }
        }

        public string FileRef
        {
            get
            {
                if (_item.FieldValues.ContainsKey("FileRef"))
                    return (string)_item.FieldValues["FileRef"];
                return null;
            }
        }

        public Guid UniqueId
        {
            get
            {
                if (_item.FieldValues.ContainsKey("UniqueId"))
                    return (Guid)_item.FieldValues["UniqueId"];
                return Guid.Empty;
            }
        }

        private SPOFile _file = null;
        public SPOFile GetFile()
        {
            if (_file != null) return _file;

            var file = _item.File;
            SPOFile.LoadFile(SPOSiteContext.CurrentSiteContext.Context, file);
            return new SPOFile(file);
        }

        public object this[string fieldName]
        {
            get
            {
                return _item[fieldName];
            }
            set
            {
                _item[fieldName] = value;
                _hasChanges = true;
            }
        }

        public void Delete()
        {
            _item.DeleteObject();
            _item.Context.ExecuteQuery();
        }

        public Guid Recycle()
        {
            var id = _item.Recycle();
            _item.Context.ExecuteQuery();
            return id.Value;
        }


        public BasePermissions GetUserEffectivePermissions(string userName)
        {
            var perm = _item.GetUserEffectivePermissions(userName);
            _item.Context.ExecuteQuery();
            return perm.Value;
        }

        public string GetWOPIFrameUrl(SPWOPIFrameAction action)
        {
            var url = _item.GetWOPIFrameUrl(action);
            _item.Context.ExecuteQuery();
            return url.Value;
        }

        public SPOContentType GetContentType()
        {
            if (_ct != null) return _ct;

            SPOContentType.LoadContentType(SPOSiteContext.CurrentSiteContext.Context, _item.ContentType);
            _ct = new SPOContentType(_item.ContentType);
            return _ct;
        }

        public Dictionary<string, object> FieldValues
        {
            get
            {
                return _item.FieldValues;
            }
        }


        private SPOList _basicParentList = null;
        private SPOList _detailedParentList = null;
        public SPOList GetParentList(bool detailed)
        {
            if (detailed && _detailedParentList != null) return _detailedParentList;
            if (!detailed && _basicParentList != null) return _basicParentList;

            if (detailed && _detailedParentList == null)
            {
                SPOList.LoadList(SPOSiteContext.CurrentSiteContext.Context, _item.ParentList, true);
                _detailedParentList = new SPOList(_item.ParentList);
                _basicParentList = _detailedParentList;
                return _detailedParentList;
            }
            else if (!detailed && _basicParentList == null)
            {
                SPOList.LoadList(SPOSiteContext.CurrentSiteContext.Context, _item.ParentList, false);
                _basicParentList = new SPOList(_item.ParentList);
                return _basicParentList;
            }
            return null;
        }

        private SPOFolder _basicFolder = null;
        private SPOFolder _detailedFolder = null;
        public SPOFolder GetFolder(bool detailed)
        {
            if (detailed && _detailedFolder != null) return _detailedFolder;
            if (!detailed && _basicFolder != null) return _basicFolder;

            if (detailed && _detailedFolder == null)
            {
                SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _item.Folder, true, true);
                _detailedFolder = new SPOFolder(_item.Folder, true);
                _basicFolder = _detailedFolder;
                return _detailedFolder;
            }
            else if (!detailed && _basicFolder == null)
            {
                SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _item.Folder, false, true);
                _basicFolder = new SPOFolder(_item.Folder, true);
                return _basicFolder;
            }
            return null;
        }


        public void Refresh()
        {
            _item.RefreshLoad();
            _item.Context.ExecuteQuery();

            _ct = null;
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _item.Update();
                _item.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

    }
}
