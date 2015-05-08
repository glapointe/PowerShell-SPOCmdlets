using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Exceptions;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOFolder : SPOClientObject
    {
        private Folder _folder = null;
        private bool _hasChanges = false;
        private bool _isListContext = false;
        public SPOFolder(Folder folder, bool isListContext)
        {
            _folder = folder;
            _isListContext = isListContext;
        }

        public Folder Folder { get { return _folder; } }

        public int ItemCount
        {
            get
            {
                if (_folder.IsPropertyAvailable("ItemCount"))
                    return _folder.ItemCount;
                return 0;
            }
        }

        public string Name
        {
            get
            {
                if (_folder.IsPropertyAvailable("Name"))
                    return _folder.Name;
                return null;
            }
        }

        public string ServerRelativeUrl
        {
            get
            {
                if (_folder.IsPropertyAvailable("ServerRelativeUrl"))
                    return _folder.ServerRelativeUrl;
                return null;
            }
        }

        public string WelcomePage
        {
            get
            {
                if (_folder.IsPropertyAvailable("WelcomePage"))
                    return _folder.WelcomePage;
                return null;
            }
            set
            {
                _folder.WelcomePage = value;
                _hasChanges = true;
            }
        }


        public IList<ContentTypeId> ContentTypeOrder
        {
            get
            {
                if (_folder.IsObjectPropertyInstantiated("ContentTypeOrder"))
                    return _folder.ContentTypeOrder;
                return null;
            }
        }

        //public FileCollection Files { get; private set; }
        //public ListItem ListItemAllFields { get; private set; }

        private SPOFolder _basicParentFolder = null;
        private SPOFolder _detailedParentFolder = null;
        public SPOFolder GetParentFolder(bool detailed)
        {
            if (detailed && _detailedParentFolder != null) return _detailedParentFolder;
            if (!detailed && _basicParentFolder != null) return _basicParentFolder;

            if (detailed && _detailedParentFolder == null)
            {
                LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _folder.ParentFolder, true, _isListContext);
                _detailedParentFolder = new SPOFolder(_folder.ParentFolder, _isListContext);
                _basicParentFolder = _detailedParentFolder;
                return _detailedParentFolder;
            }
            else if (!detailed && _basicParentFolder == null)
            {
                LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _folder.ParentFolder, false, _isListContext);
                _basicParentFolder = new SPOFolder(_folder.ParentFolder, _isListContext);
                return _basicParentFolder;
            }
            return null;
        }

        private IList<ContentTypeId> _uniqueContentTypeOrder = null;
        public IList<ContentTypeId> UniqueContentTypeOrder
        {
            get
            {
                if (_uniqueContentTypeOrder == null && _folder.IsObjectPropertyInstantiated("UniqueContentTypeOrder"))
                    _uniqueContentTypeOrder = _folder.UniqueContentTypeOrder;
                return _uniqueContentTypeOrder;
            }
            set
            {
                _folder.UniqueContentTypeOrder = _uniqueContentTypeOrder;
                _hasChanges = true;
            }
        }


        private SPOPropertyBag _properties = null;
        public SPOPropertyBag Properties
        {
            get
            {
                if (_properties == null && _folder.IsObjectPropertyInstantiated("Properties"))
                    _properties = new SPOPropertyBag(this, _folder.Properties);
                return _properties;
            }
        }

        private List<SPOFolder> _basicFolders = null;
        private List<SPOFolder> _detailedFolders = null;
        public List<SPOFolder> GetFolders(bool detailed)
        {
            if (!_folder.IsObjectPropertyInstantiated("Folders"))
            {
                _folder.Context.Load(_folder.Folders);
                _folder.Context.ExecuteQuery();
            }
            if (detailed && _detailedFolders == null)
            {
                _detailedFolders = new List<SPOFolder>();
                foreach (Folder f in _folder.Folders)
                {
                    LoadFolder(SPOSiteContext.CurrentSiteContext.Context, f, true, _isListContext);
                    _detailedFolders.Add(new SPOFolder(f, _isListContext));
                }
                _basicFolders = _detailedFolders;
                return _detailedFolders;
            }
            else if (!detailed && _basicFolders == null)
            {
                _basicFolders = new List<SPOFolder>();
                foreach (Folder f in _folder.Folders)
                {
                    LoadFolder(SPOSiteContext.CurrentSiteContext.Context, f, false, _isListContext);
                    _basicFolders.Add(new SPOFolder(f, _isListContext));
                }
                return _basicFolders;
            }
            return null;
        }

        public SPOFile AddTemplateFile(string leafName, TemplateFileType templateType)
        {
            return AddTemplateFile(leafName, templateType, false);
        }
        internal SPOFile AddTemplateFile(string leafName, TemplateFileType templateType, bool throwIfExists)
        {
            if (string.IsNullOrEmpty(leafName)) throw new ArgumentNullException("The leafName name cannot be null.");
            if (!leafName.ToLower().EndsWith(".aspx"))
                leafName += ".aspx";

            SPOFile existingFile = null;
            try
            {
                existingFile = GetFile(leafName);
            }
            catch { }
            if (existingFile != null && existingFile.Exists)
            {
                if (throwIfExists)
                    throw new FileExistsException("The specified file already exists.");

                return existingFile;
            }

            string url = ServerRelativeUrl.TrimEnd('/') + "/" + leafName;
            var file = _folder.Files.AddTemplateFile(url, templateType);
            SPOFile.LoadFile(SPOSiteContext.CurrentSiteContext.Context, file);
            return new SPOFile(file);
        }
        public SPOFile AddFile(string file, bool overwrite, string leafName, Dictionary<string, object> fieldValues)
        {
            byte[] bfile = System.IO.File.ReadAllBytes(file);
            return AddFile(bfile, overwrite, leafName, fieldValues);
        }
        public SPOFile AddFile(byte[] content, bool overwrite, string leafName, Dictionary<string, object> fieldValues)
        {
            FileCreationInformation fci = new FileCreationInformation();
            fci.Content = content;
            fci.Overwrite = overwrite;
            fci.Url = leafName;
            File file = _folder.Files.Add(fci);
            ListItem item = file.ListItemAllFields;
            if (fieldValues != null)
            {
                foreach (string fieldName in fieldValues.Keys)
                {
                    item[fieldName] = fieldValues[fieldName];
                }
            }
            item.Update();

            _folder.Context.ExecuteQuery();

            SPOFile.LoadFile(SPOSiteContext.CurrentSiteContext.Context, file);

            return new SPOFile(file);
        }

        public SPOFile GetFile(string fileName)
        {
            var file = _folder.Files.GetByUrl(fileName);
            SPOFile.LoadFile(SPOSiteContext.CurrentSiteContext.Context, file);
            return new SPOFile(file);
        }


        public void Refresh()
        {
            _folder.RefreshLoad();
            _folder.Context.ExecuteQuery();
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _folder.Update();
                _folder.Context.ExecuteQuery();
                _hasChanges = false;
            }
            if (_properties != null)
                _properties.Update();
        }

        internal static void LoadFolder(CmdletContext ctx, Folder folder, bool detailed, bool isListContext)
        {
            ctx.Load(folder);
            ctx.Load(folder, f => f.Name, f => f.ServerRelativeUrl);
            if (detailed)
            {
                if (isListContext)
                {
                    ctx.Load(folder,
                        f => f.ContentTypeOrder,
                        f => f.Folders,
                        f => f.Properties,
                        f => f.UniqueContentTypeOrder);
                }
                else
                {
                    ctx.Load(folder,
                        f => f.Folders,
                        f => f.Properties);
                }
            }
            ctx.ExecuteQuery();
            if (!isListContext && detailed)
            {
                // There are some cases where we don't know if it's a list context until it's loaded.
                if (folder.Properties.FieldValues.ContainsKey("vti_listname") && folder.Properties["vti_listname"] != null)
                {
                    ctx.Load(folder,
                        f => f.ContentTypeOrder,
                        f => f.Folders,
                        f => f.Properties,
                        f => f.UniqueContentTypeOrder);
                    ctx.ExecuteQuery();
                }
            }
        }
    }
}
