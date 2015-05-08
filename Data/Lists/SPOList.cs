using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;
using Lapointe.SharePointOnline.PowerShell.Data.EventReceivers;
using Lapointe.SharePointOnline.PowerShell.Data.Fields;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOList : SPOClientObject
    {
        private List _list = null;
        private bool _hasChanges = false;

        public SPOList(List list)
        {
            _list = list;

            //if (list.IsPropertyAvailable("EffectiveBasePermissions"))
            //    EffectiveBasePermissions = list.EffectiveBasePermissions;

            //if (list.IsPropertyAvailable("EffectiveBasePermissionsForUI"))
            //    EffectiveBasePermissionsForUI = list.EffectiveBasePermissionsForUI;

            //if (list.IsObjectPropertyInstantiated("EventReceivers"))
            //    EventReceivers = list.EventReceivers;

            //if (list.IsObjectPropertyInstantiated("Fields"))
            //    Fields = list.Fields;

            //if (list.IsObjectPropertyInstantiated("Forms"))
            //    Forms = list.Forms;

            //if (list.IsObjectPropertyInstantiated("ParentWeb"))
            //    ParentWeb = list.ParentWeb;

            //if (list.IsPropertyAvailable("DataSource"))
            //    DataSource = list.DataSource;
        }

        public List List { get { return _list; } }



        public bool AllowContentTypes
        {
            get
            {
                if (_list.IsPropertyAvailable("AllowContentTypes"))
                    return _list.AllowContentTypes;
                return false;
            }
        }

        public int BaseTemplate
        {
            get
            {
                if (_list.IsPropertyAvailable("BaseTemplate"))
                    return _list.BaseTemplate;
                return 0;
            }
        }
        public BaseType BaseType
        {
            get
            {
                if (_list.IsPropertyAvailable("BaseType"))
                    return _list.BaseType;
                return Microsoft.SharePoint.Client.BaseType.None;
            }
        }
        public BrowserFileHandling BrowserFileHandling
        {
            get
            {
                if (_list.IsPropertyAvailable("BrowserFileHandling"))
                    return _list.BrowserFileHandling;
                return Microsoft.SharePoint.Client.BrowserFileHandling.Permissive;
            }
        }
        public bool ContentTypesEnabled
        {
            get
            {
                if (_list.IsPropertyAvailable("ContentTypesEnabled"))
                    return _list.ContentTypesEnabled;
                return false;
            }
            set
            {
                _list.ContentTypesEnabled = value;
                _hasChanges = true;
            }
        }
        public DateTime Created
        {
            get
            {
                if (_list.IsPropertyAvailable("Created"))
                    return _list.Created;
                return DateTime.MinValue;
            }
        }
        //public ListDataSource DataSource { get; private set; }
        public Guid DefaultContentApprovalWorkflowId
        {
            get
            {
                if (_list.IsPropertyAvailable("DefaultContentApprovalWorkflowId"))
                    return _list.DefaultContentApprovalWorkflowId;
                return Guid.Empty;
            }
            set
            {
                _list.DefaultContentApprovalWorkflowId = value;
                _hasChanges = true;
            }
        }
        public string DefaultDisplayFormUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("DefaultDisplayFormUrl"))
                    return _list.DefaultDisplayFormUrl;
                return null;
            }
            set
            {
                _list.DefaultDisplayFormUrl = value;
                _hasChanges = true;
            }
        }
        public string DefaultEditFormUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("DefaultEditFormUrl"))
                    return _list.DefaultEditFormUrl;
                return null;
            }
            set
            {
                _list.DefaultEditFormUrl = value;
                _hasChanges = true;
            }
        }
        public string DefaultNewFormUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("DefaultNewFormUrl"))
                    return _list.DefaultNewFormUrl;
                return null;
            }
            set
            {
                _list.DefaultNewFormUrl = value;
                _hasChanges = true;
            }
        }
        public string DefaultViewUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("DefaultViewUrl"))
                    return _list.DefaultViewUrl;
                return null;
            }
        }
        public string Description
        {
            get
            {
                if (_list.IsPropertyAvailable("Description"))
                    return _list.Description;
                return null;
            }
            set
            {
                _list.Description = value;
                _hasChanges = true;
            }
        }
        public string Direction
        {
            get
            {
                if (_list.IsPropertyAvailable("Direction"))
                    return _list.Direction;
                return null;
            }
            set
            {
                _list.Direction = value;
                _hasChanges = true;
            }
        }
        public string DocumentTemplateUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("DocumentTemplateUrl"))
                    return _list.DocumentTemplateUrl;
                return null;
            }
            set
            {
                _list.DocumentTemplateUrl = value;
                _hasChanges = true;
            }
        }
        public DraftVisibilityType DraftVersionVisibility
        {
            get
            {
                if (_list.IsPropertyAvailable("DraftVersionVisibility"))
                    return _list.DraftVersionVisibility;
                return DraftVisibilityType.Reader;
            }
            set
            {
                _list.DraftVersionVisibility = value;
                _hasChanges = true;
            }
        }
        //public BasePermissions EffectiveBasePermissions { get; private set; }
        //public BasePermissions EffectiveBasePermissionsForUI { get; private set; }
        public bool EnableAttachments
        {
            get
            {
                if (_list.IsPropertyAvailable("EnableAttachments"))
                    return _list.EnableAttachments;
                return false;
            }
            set
            {
                _list.EnableAttachments = value;
                _hasChanges = true;
            }
        }
        public bool EnableFolderCreation
        {
            get
            {
                if (_list.IsPropertyAvailable("EnableFolderCreation"))
                    return _list.EnableFolderCreation;
                return false;
            }
            set
            {
                _list.EnableFolderCreation = value;
                _hasChanges = true;
            }
        }
        public bool EnableMinorVersions
        {
            get
            {
                if (_list.IsPropertyAvailable("EnableMinorVersions"))
                    return _list.EnableMinorVersions;
                return false;
            }
            set
            {
                _list.EnableMinorVersions = value;
                _hasChanges = true;
            }
        }
        public bool EnableModeration
        {
            get
            {
                if (_list.IsPropertyAvailable("EnableModeration"))
                    return _list.EnableModeration;
                return false;
            }
            set
            {
                _list.EnableModeration = value;
                _hasChanges = true;
            }
        }
        public bool EnableVersioning
        {
            get
            {
                if (_list.IsPropertyAvailable("EnableVersioning"))
                    return _list.EnableVersioning;
                return false;
            }
            set
            {
                _list.EnableVersioning = value;
                _hasChanges = true;
            }
        }
        public string EntityTypeName
        {
            get
            {
                if (_list.IsPropertyAvailable("EntityTypeName"))
                    return _list.EntityTypeName;
                return null;
            }
        }
        //public EventReceiverDefinitionCollection EventReceivers { get; private set; }
        //public FieldCollection Fields { get; private set; }
        public bool ForceCheckout
        {
            get
            {
                if (_list.IsPropertyAvailable("ForceCheckout"))
                    return _list.ForceCheckout;
                return false;
            }
            set
            {
                _list.ForceCheckout = value;
                _hasChanges = true;
            }
        }
        //public FormCollection Forms { get; private set; }
        public bool HasExternalDataSource
        {
            get
            {
                if (_list.IsPropertyAvailable("HasExternalDataSource"))
                    return _list.HasExternalDataSource;
                return false;
            }
        }
        public bool HasUniqueRoleAssignments
        {
            get
            {
                if (_list.IsPropertyAvailable("HasUniqueRoleAssignments"))
                    return _list.HasUniqueRoleAssignments;
                return false;
            }
        }
        public bool Hidden
        {
            get
            {
                if (_list.IsPropertyAvailable("Hidden"))
                    return _list.Hidden;
                return false;
            }
            set
            {
                _list.Hidden = value;
                _hasChanges = true;
            }
        }
        public Guid Id
        {
            get
            {
                if (_list.IsPropertyAvailable("Id"))
                    return _list.Id;
                return Guid.Empty;
            }
        }
        public string ImageUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("ImageUrl"))
                    return _list.ImageUrl;
                return null;
            }
            set
            {
                _list.ImageUrl = value;
                _hasChanges = true;
            }
        }

        private SPOInformationRightsManagementSettings _informationRightsManagementSettings = null;
        public SPOInformationRightsManagementSettings GetInformationRightsManagementSettings()
        {
            if (_informationRightsManagementSettings != null || this.BaseTemplate == 109) return _informationRightsManagementSettings;

            _list.Context.Load(_list, l => l.InformationRightsManagementSettings);
            _list.Context.ExecuteQuery();
            _informationRightsManagementSettings = new SPOInformationRightsManagementSettings(_list.InformationRightsManagementSettings, _list);

            return _informationRightsManagementSettings;
        }
        public bool IrmEnabled
        {
            get
            {
                if (_list.IsPropertyAvailable("IrmEnabled"))
                    return _list.IrmEnabled;
                return false;
            }
            set
            {
                _list.IrmEnabled = value;
                _hasChanges = true;
            }
        }
        public bool IrmExpire
        {
            get
            {
                if (_list.IsPropertyAvailable("IrmExpire"))
                    return _list.IrmExpire;
                return false;
            }
            set
            {
                _list.IrmExpire = value;
                _hasChanges = true;
            }
        }
        public bool IrmReject
        {
            get
            {
                if (_list.IsPropertyAvailable("IrmReject"))
                    return _list.IrmReject;
                return false;
            }
            set
            {
                _list.IrmReject = value;
                _hasChanges = true;
            }
        }
        public bool IsApplicationList
        {
            get
            {
                if (_list.IsPropertyAvailable("IsApplicationList"))
                    return _list.IsApplicationList;
                return false;
            }
            set
            {
                _list.IsApplicationList = value;
                _hasChanges = true;
            }
        }
        public bool IsCatalog
        {
            get
            {
                if (_list.IsPropertyAvailable("IsCatalog"))
                    return _list.IsCatalog;
                return false;
            }
        }
        public bool IsPrivate
        {
            get
            {
                if (_list.IsPropertyAvailable("IsPrivate"))
                    return _list.IsPrivate;
                return false;
            }
        }
        public bool IsSiteAssetsLibrary
        {
            get
            {
                if (_list.IsPropertyAvailable("IsSiteAssetsLibrary"))
                    return _list.IsSiteAssetsLibrary;
                return false;
            }
        }
        public int ItemCount
        {
            get
            {
                if (_list.IsPropertyAvailable("ItemCount"))
                    return _list.ItemCount;
                return 0;
            }
        }
        public DateTime LastItemDeletedDate
        {
            get
            {
                if (_list.IsPropertyAvailable("LastItemDeletedDate"))
                    return _list.LastItemDeletedDate;
                return DateTime.MinValue;
            }
        }
        public DateTime LastItemModifiedDate
        {
            get
            {
                if (_list.IsPropertyAvailable("LastItemModifiedDate"))
                    return _list.LastItemModifiedDate;
                return DateTime.MinValue;
            }
            set
            {
                _list.LastItemModifiedDate = value;
                _hasChanges = true;
            }
        }
        public string ListItemEntityTypeFullName
        {
            get
            {
                if (_list.IsPropertyAvailable("ListItemEntityTypeFullName"))
                    return _list.ListItemEntityTypeFullName;
                return null;
            }
        }
        public bool MultipleDataList
        {
            get
            {
                if (_list.IsPropertyAvailable("MultipleDataList"))
                    return _list.MultipleDataList;
                return false;
            }
            set
            {
                _list.MultipleDataList = value;
                _hasChanges = true;
            }
        }
        public bool NoCrawl
        {
            get
            {
                if (_list.IsPropertyAvailable("NoCrawl"))
                    return _list.NoCrawl;
                return false;
            }
            set
            {
                _list.NoCrawl = value;
                _hasChanges = true;
            }
        }
        public bool OnQuickLaunch
        {
            get
            {
                if (_list.IsPropertyAvailable("OnQuickLaunch"))
                    return _list.OnQuickLaunch;
                return false;
            }
            set
            {
                _list.OnQuickLaunch = value;
                _hasChanges = true;
            }
        }
        //public Web ParentWeb { get; private set; }
        public string ParentWebUrl
        {
            get
            {
                if (_list.IsPropertyAvailable("ParentWebUrl"))
                    return _list.ParentWebUrl;
                return null;
            }
        }

        public string SchemaXml
        {
            get
            {
                if (_list.IsPropertyAvailable("SchemaXml"))
                    return _list.SchemaXml;
                return null;
            }
        }
        public bool ServerTemplateCanCreateFolders
        {
            get
            {
                if (_list.IsPropertyAvailable("ServerTemplateCanCreateFolders"))
                    return _list.ServerTemplateCanCreateFolders;
                return false;
            }
        }
        public Guid TemplateFeatureId
        {
            get
            {
                if (_list.IsPropertyAvailable("TemplateFeatureId"))
                    return _list.TemplateFeatureId;
                return Guid.Empty;
            }
        }
        public string Title
        {
            get
            {
                if (_list.IsPropertyAvailable("Title"))
                    return _list.Title;
                return null;
            }
            set
            {
                _list.Title = value;
                _hasChanges = true;
            }
        }
        public string Url
        {
            get
            {
                SPOFolder f = GetRootFolder(false);
                return f.ServerRelativeUrl;
            }
        }

        public UserCustomActionCollection UserCustomActions
        {
            get
            {
                if (_list.IsObjectPropertyInstantiated("UserCustomActions"))
                    return _list.UserCustomActions;
                return null;
            }
        }
        public string ValidationFormula
        {
            get
            {
                if (_list.IsPropertyAvailable("ValidationFormula"))
                    return _list.ValidationFormula;
                return null;
            }
            set
            {
                _list.ValidationFormula = value;
                _hasChanges = true;
            }
        }
        public string ValidationMessage
        {
            get
            {
                if (_list.IsPropertyAvailable("ValidationMessage"))
                    return _list.ValidationMessage;
                return null;
            }
            set
            {
                _list.ValidationMessage = value;
                _hasChanges = true;
            }
        }
        public WorkflowAssociationCollection WorkflowAssociations
        {
            get
            {
                if (_list.IsObjectPropertyInstantiated("WorkflowAssociations"))
                    return _list.WorkflowAssociations;
                return null;
            }
        }

        private SPOFolder _basicRootFolder = null;
        private SPOFolder _detailedRootFolder = null;
        public SPOFolder GetRootFolder(bool detailed)
        {
            if (detailed && _detailedRootFolder != null) return _detailedRootFolder;
            if (!detailed && _basicRootFolder != null) return _basicRootFolder;

            if (detailed && _detailedRootFolder == null)
            {
                SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _list.RootFolder, true, true);
                _detailedRootFolder = new SPOFolder(_list.RootFolder, true);
                _basicRootFolder = _detailedRootFolder;
                return _detailedRootFolder;
            }
            else if (!detailed && _basicRootFolder == null)
            {
                SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, _list.RootFolder, false, true);
                _basicRootFolder = new SPOFolder(_list.RootFolder, true);
                return _basicRootFolder;
            }
            return null;
        }

        private SPOView _defaultView = null;
        public SPOView GetDefaultView()
        {
            if (_defaultView != null) return _defaultView;

            var defaultView = _list.DefaultView;
            SPOView.LoadView(SPOSiteContext.CurrentSiteContext.Context, defaultView);
            _defaultView = new SPOView(defaultView, _list.Views);
            return _defaultView;
        }

        private List<SPOView> _views = null;
        public List<SPOView> GetViews()
        {
            if (_views != null) return _views;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var views = _list.Views;
            ctx.Load(views);
            ctx.ExecuteQuery();
            _views = new List<SPOView>();
            foreach (View view in views)
            {
                SPOView.LoadView(ctx, view);
                _views.Add(new SPOView(view, _list.Views));
            }
            return _views;
        }

        private List<SPOContentType> _contentTypes = null;
        public List<SPOContentType> GetContentTypes()
        {
            if (_contentTypes != null) return _contentTypes;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var contentTypes = List.ContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();
            _contentTypes = new List<SPOContentType>();
            foreach (ContentType contentType in contentTypes)
            {
                SPOContentType.LoadContentType(ctx, contentType);
                _contentTypes.Add(new SPOContentType(contentType));
            }
            return _contentTypes;
        }


        private List<SPOField> _fields = null;
        public List<SPOField> GetFields()
        {
            if (_fields != null) return _fields;

            var fields = _list.Fields;
            _list.Context.Load(fields);
            _list.Context.ExecuteQuery();
            _fields = new List<SPOField>();
            foreach (Field f in fields)
            {
                _fields.Add(new SPOField(f));
            }
            return _fields;
        }
        public SPOField AddFieldAsXml(string schemaXml, bool addToDefaultView, AddFieldOptions options)
        {
            var field = _list.Fields.AddFieldAsXml(schemaXml, addToDefaultView, options);
            _list.Context.Load(field);
            _list.Context.ExecuteQuery();
            return new SPOField(field);
        }
        public SPOField AddField(SPOField existingField)
        {
            return AddField(existingField.Field);
        }
        public SPOField AddField(Field existingField)
        {
            var field = _list.Fields.Add(existingField);
            _list.Context.Load(field);
            _list.Context.ExecuteQuery();
            return new SPOField(field);
        }
        public SPOField GetFieldByInternalNameOrTitle(string strName)
        {
            var field = _list.Fields.GetByInternalNameOrTitle(strName);
            _list.Context.Load(field);
            _list.Context.ExecuteQuery();
            return new SPOField(field);
        }
        public SPOField GetFieldByTitle(string title)
        {
            var field = _list.Fields.GetByTitle(title);
            _list.Context.Load(field);
            _list.Context.ExecuteQuery();
            return new SPOField(field);
        }

        private List<SPOEventReceiverDefinition> _eventReceivers = null;
        public List<SPOEventReceiverDefinition> GetEventReceivers()
        {
            if (_eventReceivers != null) return _eventReceivers;

            var eventReceivers = _list.EventReceivers;
            _list.Context.Load(eventReceivers);
            _list.Context.ExecuteQuery();
            _eventReceivers = new List<SPOEventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in eventReceivers)
            {
                _eventReceivers.Add(new SPOEventReceiverDefinition(erd, _list.EventReceivers));
            }
            return _eventReceivers;
        }

        public SPOEventReceiverDefinition AddEventReceiver(EventReceiverDefinitionCreationInformation erdci)
        {
            SPOEventReceiverDefinition erd = new SPOEventReceiverDefinition(_list.EventReceivers.Add(erdci), _list.EventReceivers);
            _list.EventReceivers.Context.ExecuteQuery();
            return erd;
        }
        public SPOEventReceiverDefinition AddEventReceiver(
            EventReceiverType eventType, 
            string receiverAssembly, 
            string receiverClass,
            string receiverName,
            string receiverUrl,
            int sequenceNumber,
            EventReceiverSynchronization synchronization)
        {
            EventReceiverDefinitionCreationInformation erdci = new EventReceiverDefinitionCreationInformation
            {
                EventType = eventType,
                ReceiverAssembly = receiverAssembly,
                ReceiverClass = receiverClass,
                ReceiverName = receiverName,
                ReceiverUrl = receiverUrl,
                SequenceNumber = sequenceNumber,
                Synchronization = synchronization
            };
            return AddEventReceiver(erdci);
        }


        public void BreakRoleInheritance(bool copyRoleAssignments, bool clearScopes)
        {
            _list.BreakRoleInheritance(copyRoleAssignments, clearScopes);
            _list.Context.ExecuteQuery();
            Refresh();
        }

        public void ResetRoleInheritance()
        {
            _list.ResetRoleInheritance();
            _list.Context.ExecuteQuery();
            Refresh();
        }


        public SPOListItem AddItem(Hashtable fieldValues)
        {
            return AddItem(null, fieldValues);
        }
        public SPOListItem AddItem(string listUrlFolderPath, Hashtable fieldValues)
        {
            Hashtable hash = fieldValues.Clone() as Hashtable;
            var properties = new Dictionary<string, object>();
            foreach (var key in hash.Keys)
            {
                string keyValue = key.ToString();
                properties.Add(keyValue, hash[key]);
            }

            return AddItem(listUrlFolderPath, properties);
        }
        public SPOListItem AddItem(Dictionary<string, object> fieldValues)
        {
            return AddItem(null, fieldValues);
        }
        public SPOListItem AddItem(string listUrlFolderPath, Dictionary<string, object> fieldValues)
        {
            ListItemCreationInformation lici = null;
            if (!string.IsNullOrEmpty(listUrlFolderPath))
            {
                SPOFolder rootFolder = GetRootFolder(false);
                lici = new ListItemCreationInformation();
                lici.FolderUrl =  rootFolder.ServerRelativeUrl.TrimEnd('/') + "/" + listUrlFolderPath;
            }
            ListItem item = _list.AddItem(lici);
            if (fieldValues != null)
            {
                foreach (string fieldName in fieldValues.Keys)
                {
                    item[fieldName] = fieldValues[fieldName];
                }
            }

            item.Update();
            item.Context.Load(item, i => i.Id);
            item.Context.ExecuteQuery();
            return GetItemById(item.Id);
        }

        public SPOFile AddFile(string file, bool overwrite, string leafName, Dictionary<string, object> fieldValues)
        {
            byte[] bfile = System.IO.File.ReadAllBytes(file);
            return AddFile(bfile, overwrite, leafName, fieldValues);
        }
        public SPOFile AddFile(byte[] content, bool overwrite, string leafName, Dictionary<string, object> fieldValues)
        {
            SPOFolder rootFolder = GetRootFolder(false);
            return rootFolder.AddFile(content, overwrite, leafName, fieldValues);
        }

        HybridDictionary _items = new HybridDictionary();
        public SPOListItem GetItemById(int id)
        {
            if (_items.Contains(id)) return _items[id] as SPOListItem;
            var item = _list.GetItemById(id);
            _list.Context.Load(item);
            _list.Context.ExecuteQuery();
            SPOListItem li = new SPOListItem(item);
            _items.Add(id, li);
            return li;
        }

        public List<SPOListItem> GetItems(int rowLimit, params string[] viewFields)
        {
            CamlQuery query = CamlQuery.CreateAllItemsQuery(rowLimit, viewFields);
            return GetItems(query);
        }

        public List<SPOListItem> GetItems(string viewXml)
        {
            CamlQuery query = new CamlQuery();
            query.ViewXml = viewXml;
            return GetItems(query);
        }
        public List<SPOListItem> GetItems(CamlQuery query)
        {
            var items = _list.GetItems(query);
            _list.Context.Load(items);
            _list.Context.ExecuteQuery();

            List<SPOListItem> li = new List<SPOListItem>();
            foreach (var listItem in items)
            {
                SPOListItem li1 = new SPOListItem(listItem);
                li.Add(li1);
            }
            return li;
        }
        
        public SPOFolder GetFolder(string listRelativeFolderUrl, bool detailed, bool createIfMissing)
        {
            SPOFolder rootFolder = this.GetRootFolder(detailed);
            if (string.IsNullOrEmpty(listRelativeFolderUrl))
                return rootFolder;

            string rootFolderUrl = rootFolder.ServerRelativeUrl.TrimEnd('/');
            Folder folder = null;
            try
            {
                folder = _list.ParentWeb.GetFolderByServerRelativeUrl(rootFolderUrl + "/" + listRelativeFolderUrl);
                SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, folder, detailed, true);
            }
            catch (ServerException)
            {
                folder = null;
            }

            if (folder == null && createIfMissing)
            {
                if (!this.EnableFolderCreation)
                {
                    this.EnableFolderCreation = true;
                    this.Update();
                }

                string[] folders = listRelativeFolderUrl.Trim('/').Split('/');

                string folderPath = string.Empty;
                for (int i = 0; i < folders.Length; i++)
                {
                    string parentFolder = folderPath;
                    folderPath += "/" + folders[i];
                    folder = null;
                    try
                    {
                        folder = _list.ParentWeb.GetFolderByServerRelativeUrl(rootFolderUrl + folderPath);
                        SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, folder, detailed, true);
                    }
                    catch (ServerException) 
                    {
                        folder = null;
                    }
                    if (folder == null)
                    {
                        ListItemCreationInformation lici = new ListItemCreationInformation();
                        lici.LeafName = folders[i];
                        lici.FolderUrl = rootFolderUrl + parentFolder; // folderPath.Trim('/');
                        lici.UnderlyingObjectType = FileSystemObjectType.Folder;

                        var item = _list.AddItem(lici);
                        item["Title"] = folders[i];
                        item.Update();
                        item.Context.ExecuteQuery();
                        folder = item.Folder;
                        SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, folder, detailed, true);
                    }
                }
            }
            // Still no folder so error out
            if (folder == null)
                throw new Exception(string.Format("The folder '{0}' could not be found.", listRelativeFolderUrl));
            return new SPOFolder(folder, true);
        }
         

        public void Delete()
        {
            _list.DeleteObject();
            _list.Context.ExecuteQuery();
        }

        public Guid Recycle()
        {
            var id = _list.Recycle();
            _list.Context.ExecuteQuery();
            return id.Value;
        }


        public void Refresh()
        {
            _list.RefreshLoad();
            _list.Context.ExecuteQuery();
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _list.Update();
                _list.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

        internal static void LoadList(CmdletContext ctx, List list, bool detailed)
        {
            ctx.Load(list);
            ctx.Load(list,
                l => l.RootFolder,
                l => l.DefaultViewUrl,
                l => l.HasUniqueRoleAssignments);
            if (detailed)
            {
                ctx.Load(list,
                    l => l.DefaultDisplayFormUrl,
                    l => l.DefaultEditFormUrl,
                    l => l.DefaultNewFormUrl,
                    l => l.IsSiteAssetsLibrary,
                    l => l.OnQuickLaunch,
                    l => l.SchemaXml,
                    l => l.UserCustomActions,
                    l => l.ValidationFormula,
                    l => l.ValidationMessage,
                    l => l.WorkflowAssociations);
                SPOFolder.LoadFolder(ctx, list.RootFolder, detailed, true);
            }
            else
                ctx.ExecuteQuery();
        }

    }
}
