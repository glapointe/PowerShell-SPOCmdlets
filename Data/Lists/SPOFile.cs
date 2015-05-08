using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Principals;
using Lapointe.SharePointOnline.PowerShell.Data.WebParts;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.WebParts;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOFile : SPOClientObject
    {
        protected File _file = null;

        public SPOFile(File file)
        {
            _file = file;
        }

        public File File { get { return _file; } }

        private SPOUser _author = null;
        public SPOUser Author
        {
            get
            {
                if (_author == null && _file.IsObjectPropertyInstantiated("Author"))
                    _author = new SPOUser(_file.Author);
                return _author;
            }
        }

        private SPOUser _checkedOutByUser = null;
        public SPOUser CheckedOutByUser
        {
            get
            {
                if (_checkedOutByUser == null && _file.IsObjectPropertyInstantiated("CheckedOutByUser"))
                    _checkedOutByUser = new SPOUser(_file.CheckedOutByUser);
                return _checkedOutByUser;
            }
        }
        
        public string CheckInComment
        {
            get
            {
                if (_file.IsPropertyAvailable("CheckInComment"))
                    return _file.CheckInComment;
                return null;
            }
        }

        public CheckOutType CheckOutType
        {
            get
            {
                if (_file.IsPropertyAvailable("CheckOutType"))
                    return _file.CheckOutType;
                return CheckOutType.None;
            }
        }

        public string ContentTag
        {
            get
            {
                if (_file.IsPropertyAvailable("ContentTag"))
                    return _file.ContentTag;
                return null;
            }
        }

        public CustomizedPageStatus CustomizedPageStatus
        {
            get
            {
                if (_file.IsPropertyAvailable("CustomizedPageStatus"))
                    return _file.CustomizedPageStatus;
                return CustomizedPageStatus.None;
            }
        }

        public string ETag
        {
            get
            {
                if (_file.IsPropertyAvailable("ETag"))
                    return _file.ETag;
                return null;
            }
        }

        public bool Exists
        {
            get
            {
                if (_file.IsPropertyAvailable("Exists"))
                    return _file.Exists;
                return false;
            }
        }

        public long Length
        {
            get
            {
                if (_file.IsPropertyAvailable("Length"))
                    return _file.Length;
                return 0;
            }
        }

        public FileLevel Level
        {
            get
            {
                if (_file.IsPropertyAvailable("Level"))
                    return _file.Level;
                return FileLevel.Published;
            }
        }

        private SPOListItem _listItemAllFields = null;
        public SPOListItem ListItemAllFields
        {
            get
            {
                if (_listItemAllFields == null && _file.IsObjectPropertyInstantiated("ListItemAllFields"))
                    _listItemAllFields = new SPOListItem(_file.ListItemAllFields);
                return _listItemAllFields;
            }
        }

        private SPOUser _lockedByUser = null;
        public SPOUser LockedByUser
        {
            get
            {
                if (_lockedByUser == null && _file.IsObjectPropertyInstantiated("LockedByUser"))
                    _lockedByUser = new SPOUser(_file.LockedByUser);
                return _lockedByUser;
            }
        }

        public int MajorVersion
        {
            get
            {
                if (_file.IsPropertyAvailable("MajorVersion"))
                    return _file.MajorVersion;
                return 0;
            }
        }

        public int MinorVersion
        {
            get
            {
                if (_file.IsPropertyAvailable("MinorVersion"))
                    return _file.MinorVersion;
                return 0;
            }
        }

        private SPOUser _modifiedBy = null;
        public SPOUser ModifiedBy
        {
            get
            {
                if (_modifiedBy == null && _file.IsObjectPropertyInstantiated("ModifiedBy"))
                    _modifiedBy = new SPOUser(_file.ModifiedBy);
                return _modifiedBy;
            }
        }

        public string Name
        {
            get
            {
                if (_file.IsPropertyAvailable("Name"))
                    return _file.Name;
                return null;
            }
        }

        public string ServerRelativeUrl
        {
            get
            {
                if (_file.IsPropertyAvailable("ServerRelativeUrl"))
                    return _file.ServerRelativeUrl;
                return null;
            }
        }

        public DateTime TimeCreated
        {
            get
            {
                if (_file.IsPropertyAvailable("TimeCreated"))
                    return _file.TimeCreated;
                return DateTime.MinValue;
            }
        }

        public DateTime TimeLastModified
        {
            get
            {
                if (_file.IsPropertyAvailable("TimeLastModified"))
                    return _file.TimeLastModified;
                return DateTime.MinValue;
            }
        }

        public string Title
        {
            get
            {
                if (_file.IsPropertyAvailable("Title"))
                    return _file.Title;
                return null;
            }
        }

        public int UIVersion
        {
            get
            {
                if (_file.IsPropertyAvailable("UIVersion"))
                    return _file.UIVersion;
                return 0;
            }
        }

        public string UIVersionLabel
        {
            get
            {
                if (_file.IsPropertyAvailable("UIVersionLabel"))
                    return _file.UIVersionLabel;
                return null;
            }
        }

        //public FileVersionCollection Versions { get; }

        public void SaveFileLocally(string directory, bool overwrite)
        {
            string path = System.IO.Path.Combine(directory, Name);
            if (System.IO.File.Exists(path))
            {
                if (!overwrite) throw new Exception("The file already exists locally.");
                System.IO.File.Delete(path);
            }
            var stream = _file.OpenBinaryStream();
            _file.Context.ExecuteQuery();
            using (System.IO.Stream s = stream.Value)
            using (System.IO.FileStream fs = System.IO.File.Create(path))
            {
                s.CopyTo(fs);
            }
            Refresh();
        }

        public void SaveFileLocallyViaREST(string directory, bool overwrite)
        {
            string path = System.IO.Path.Combine(directory, Name);
            if (System.IO.File.Exists(path))
            {
                if (!overwrite) throw new Exception("The file already exists locally.");
                System.IO.File.Delete(path);
            }
            string folderUrl = ServerRelativeUrl.Substring(0, ServerRelativeUrl.Length - Name.Length - 1);
            string restCmd = string.Format("/_api/web/GetFolderByServerRelativeUrl('{0}')/files('{1}')/$value", folderUrl, Name);

            Uri uri = new Uri(_file.Context.Url.TrimEnd('/') + restCmd);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Credentials = SPOSiteContext.CurrentSiteContext.Context.Credentials;
            request.Method = WebRequestMethods.Http.Get;
            request.Headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.ContentLength > 0) response.ContentLength = response.ContentLength;
            using (System.IO.Stream s = response.GetResponseStream())
            using (System.IO.FileStream fs = System.IO.File.Create(path))
            {
                s.CopyTo(fs);
            }
        }


        public void UploadReplacementFile(byte[] content, bool checkRequiredFields, Dictionary<string, object> fieldValues)
        {
            FileSaveBinaryInformation fsbi = new FileSaveBinaryInformation();
            fsbi.CheckRequiredFields = checkRequiredFields;
            fsbi.FieldValues = fieldValues;
            fsbi.Content = content;
            fsbi.ETag = _file.ETag;
            _file.SaveBinary(fsbi);
            _file.Context.ExecuteQuery();
            Refresh();
        }

        public void Publish(string comment)
        {
            _file.Publish(comment);
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }

        public void UnPublish(string comment)
        {
            _file.UnPublish(comment);
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }
        public void Approve(string comment)
        {
            _file.Approve(comment);
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }
        public void Deny(string comment)
        {
            _file.Deny(comment);
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }

        public void CheckOut()
        {
            _file.CheckOut();
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }

        public void CheckIn(string comment, CheckinType checkinType)
        {
            _file.CheckIn(comment, checkinType);
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }

        public void UndoCheckOut()
        {
            _file.UndoCheckOut();
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();
        }

        public void Delete()
        {
            _file.DeleteObject();
            _file.Context.ExecuteQuery();
        }

        public Guid Recycle()
        {
            var id = _file.Recycle();
            _file.Context.ExecuteQuery();
            return id.Value;
        }

        public void Refresh()
        {
            _file.RefreshLoad();
            _file.Context.ExecuteQuery();

            _listItemAllFields = null;
        }

        private List<SPOWebPartDefinition> _webParts = null;
        public List<SPOWebPartDefinition> GetWebPartDefinitions()
        {
            if (_webParts != null) return _webParts;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            LimitedWebPartManager limitedWebPartManager = File.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webParts = limitedWebPartManager.WebParts;
            ctx.Load(webParts);
            ctx.ExecuteQuery();
            _webParts = new List<SPOWebPartDefinition>();
            foreach (WebPartDefinition webPartDef in webParts)
            {
                File.Context.Load(webPartDef, w => w.Id, w => w.WebPart);
                _webParts.Add(new SPOWebPartDefinition(webPartDef));
            }
            return _webParts;
        }

        public SPOWebPartDefinition AddWebPartToWebPartPage(string webPartXml, string zone, int index, string title)
        {
            LimitedWebPartManager limitedWebPartManager = File.GetLimitedWebPartManager(PersonalizationScope.Shared);
            WebPartDefinition webPartDefinition = limitedWebPartManager.ImportWebPart(webPartXml);

            var wpd = limitedWebPartManager.AddWebPart(webPartDefinition.WebPart, zone, index);
            File.Context.Load(wpd, w => w.Id, w => w.WebPart);
            File.Context.ExecuteQuery();

            var spowpd = new SPOWebPartDefinition(wpd);
            if (spowpd.WebPart.Title != title)
            {
                spowpd.WebPart.Title = title;
                spowpd.SaveWebPartChanges();
            }
            return spowpd;
        }

        internal static void LoadFile(CmdletContext ctx, File file)
        {
            ctx.Load(file);
            ctx.Load(file,
                f => f.Author,
                f => f.ModifiedBy,
                f => f.CheckedOutByUser,
                f => f.LockedByUser,
                f => f.Versions);
            ctx.ExecuteQuery();

            try
            {
                ctx.Load(file, f => f.ListItemAllFields);
                ctx.ExecuteQuery();
            }
            catch { }
        }
    }
}
