using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Lapointe.SharePointOnline.PowerShell.Data.Principals;
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.EventReceivers;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Lapointe.SharePointOnline.PowerShell.Data.Navigation;

namespace Lapointe.SharePointOnline.PowerShell.Data.Sites
{
    public class SPOWeb : SPOClientObject
    {
        private Web _web = null;
        private bool _hasChanges = false;
        private List<SPOList> _detailedLists = null;
        private List<SPOList> _lists = null;
        private List<SPOContentType> _availableContentTypes = null;
        private List<SPOContentType> _contentTypes = null;
        private List<SPOFolder> _folders = null;
        private SPOPropertyBag _allProperties = null;
        private SPOGroup _associatedMemberGroup = null;
        private SPOGroup _associatedOwnerGroup = null;
        private SPOGroup _associatedVisitorGroup = null;
        private SPOUser _currentUser = null;
        private SPONavigation _navigation = null;
        private List<SPOWeb> _subWebsForCurrentUser = null;
        private List<SPOWeb> _detailedSubWebsForCurrentUser = null;
        private List<SPOWeb> _subWebs = null;
        private List<SPOWeb> _detailedSubWebs = null;
        private List<SPOWeb> _allSubWebs = null;
        private List<SPOWeb> _detailedAllSubWebs = null;
        private List<SPOEventReceiverDefinition> _eventReceivers = null;
        private SPOUserCustomActionCollection _userCustomActions = null;

        public SPOWeb(Web web)
        {
            _web = web;


            //if (web.IsObjectPropertyInstantiated("AvailableFields"))
            //    AvailableFields = web.AvailableFields;

            //if (web.IsPropertyAvailable("EffectiveBasePermissions"))
            //    EffectiveBasePermissions = web.EffectiveBasePermissions;

            //if (web.IsObjectPropertyInstantiated("Features"))
            //    Features = web.Features;

            //if (web.IsObjectPropertyInstantiated("Fields"))
            //    Fields = web.Fields;

            //if (web.IsObjectPropertyInstantiated("ListTemplates"))
            //    ListTemplates = web.ListTemplates;

            //if (web.IsObjectPropertyInstantiated("Navigation"))
            //    Navigation = web.Navigation;

            //if (web.IsObjectPropertyInstantiated("RootFolder"))
            //    RootFolder = web.RootFolder;

            //if (web.IsObjectPropertyInstantiated("SiteGroups"))
            //    SiteGroups = web.SiteGroups;

            //if (web.IsObjectPropertyInstantiated("SiteUserInfoList"))
            //    SiteUserInfoList = web.SiteUserInfoList;

            //if (web.IsObjectPropertyInstantiated("SiteUsers"))
            //    SiteUsers = web.SiteUsers;

        }

        public Web Web { get { return _web; } }

        public SPOSite GetSite(bool detailed)
        {
            return SPOSiteContext.CurrentSiteContext.GetSite(detailed);
        }
        public List<SPOContentType> GetAvailableContentTypes()
        {
            if (_availableContentTypes != null) return _availableContentTypes;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var contentTypes = Web.AvailableContentTypes;
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();
            _availableContentTypes = new List<SPOContentType>();
            foreach (ContentType contentType in contentTypes)
            {
                SPOContentType.LoadContentType(ctx, contentType);
                _availableContentTypes.Add(new SPOContentType(contentType));
            }
            return _availableContentTypes;
        }

        public SPOList GetList(Guid id, bool detailed)
        {
            SPOListPipeBind listPB = new SPOListPipeBind(id);
            return listPB.Read(this.Web, detailed);
        }

        public SPOList GetList(string urlOrTitle, bool detailed)
        {
            SPOListPipeBind listPB = new SPOListPipeBind(urlOrTitle);
            return listPB.Read(this.Web, detailed);
        }

        public List<SPOList> GetLists(bool detailed)
        {
            if (_detailedLists != null) return _detailedLists;
            if (_lists != null && !detailed) return _lists;

            var lists = _web.Lists;
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            var spoLists = new List<SPOList>();

            foreach (var list in lists)
            {
                SPOList.LoadList(ctx, list, detailed);
                spoLists.Add(new SPOList(list));
            }
            _lists = spoLists;
            if (detailed)
                _detailedLists = spoLists;

            return spoLists;
        }

        public List<SPOContentType> GetContentTypes()
        {
            if (_contentTypes != null) return _contentTypes;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var contentTypes = Web.ContentTypes;
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

        public List<SPOFolder> GetFolders(bool detailed)
        {
            if (_folders != null) return _folders;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var folders = Web.Folders;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            _folders = new List<SPOFolder>();
            foreach (Folder folder in folders)
            {
                SPOFolder.LoadFolder(ctx, folder, detailed, false);
                _folders.Add(new SPOFolder(folder, false));
            }
            return _folders;
        }

        public SPOFile GetFileByServerRelativeUrl(string serverRelativeUrl)
        {
            var file = _web.GetFileByServerRelativeUrl(serverRelativeUrl);
            SPOFile.LoadFile(SPOSiteContext.CurrentSiteContext.Context, file);
            var spoFile = new SPOFile(file);
            if (spoFile.ListItemAllFields != null && (spoFile.ListItemAllFields.FieldValues.ContainsKey("WikiField") || 
                (serverRelativeUrl.ToLower().Contains("/sitepages/") && serverRelativeUrl.ToLower().Contains(".aspx"))))
                return new SPOWikiPage(file);
            return spoFile;
        }
        public SPOFolder GetFolderByServerRelativeUrl(string serverRelativeUrl)
        {
            var folder = _web.GetFolderByServerRelativeUrl(serverRelativeUrl);
            SPOFolder.LoadFolder(SPOSiteContext.CurrentSiteContext.Context, folder, true, false);
            return new SPOFolder(folder, false);
        }

        public bool AllowDesignerForCurrentUser
        {
            get
            {
                if (_web.IsPropertyAvailable("AllowDesignerForCurrentUser"))
                    return _web.AllowDesignerForCurrentUser;
                return false;
            }
        }

        public bool AllowMasterPageEditingForCurrentUser
        {
            get
            {
                if (_web.IsPropertyAvailable("AllowMasterPageEditingForCurrentUser"))
                    return _web.AllowMasterPageEditingForCurrentUser;
                return false;
            }
        }

        public bool AllowRevertFromTemplateForCurrentUser
        {
            get
            {
                if (_web.IsPropertyAvailable("AllowRevertFromTemplateForCurrentUser"))
                    return _web.AllowRevertFromTemplateForCurrentUser;
                return false;
            }
        }

        public bool AllowRssFeeds
        {
            get
            {
                if (_web.IsPropertyAvailable("AllowRssFeeds"))
                    return _web.AllowRssFeeds;
                return false;
            }
        }

        public SPOPropertyBag AllProperties
        {
            get
            {
                if (_allProperties == null && _web.IsObjectPropertyInstantiated("AllProperties"))
                    _allProperties = new SPOPropertyBag(this, _web.AllProperties);
                return _allProperties;
            }
        }

        public Guid AppInstanceId
        {
            get
            {
                if (_web.IsPropertyAvailable("AppInstanceId"))
                    return _web.AppInstanceId;
                return Guid.Empty;
            }
        }

        public SPOGroup AssociatedMemberGroup
        {
            get
            {
                if (_associatedMemberGroup == null && _web.IsObjectPropertyInstantiated("AssociatedMemberGroup"))
                    _associatedMemberGroup = new SPOGroup(_web.AssociatedMemberGroup);
                return _associatedMemberGroup;
            }
            internal set
            {
                _associatedMemberGroup = value;
                if (value != null)
                    _web.AssociatedMemberGroup = value.Group;
                else
                    _web.AssociatedMemberGroup = null;
                _hasChanges = true;
            }
        }

        public SPOGroup AssociatedOwnerGroup
        {
            get
            {
                if (_associatedOwnerGroup == null && _web.IsObjectPropertyInstantiated("AssociatedOwnerGroup"))
                    _associatedOwnerGroup = new SPOGroup(_web.AssociatedOwnerGroup);
                return _associatedOwnerGroup;
            }
            internal set
            {
                _associatedOwnerGroup = value;
                if (value != null)
                    _web.AssociatedOwnerGroup = value.Group;
                else
                    _web.AssociatedOwnerGroup = null; 
                _hasChanges = true;
            }
        }

        public SPOGroup AssociatedVisitorGroup
        {
            get
            {
                if (_associatedVisitorGroup == null && _web.IsObjectPropertyInstantiated("AssociatedVisitorGroup"))
                   _associatedVisitorGroup = new SPOGroup(_web.AssociatedVisitorGroup);
                return _associatedVisitorGroup;
            }
            internal set
            {
                _associatedVisitorGroup = value;
                if (value != null)
                    _web.AssociatedVisitorGroup = value.Group;
                else
                    _web.AssociatedVisitorGroup = null; 
                _hasChanges = true;
            }
        }

        //public FieldCollection AvailableFields { get; private set; }
        public short Configuration
        {
            get
            {
                if (_web.IsPropertyAvailable("Configuration"))
                    return _web.Configuration;
                return 0;
            }
        }

        public DateTime Created
        {
            get
            {
                if (_web.IsPropertyAvailable("Created"))
                    return _web.Created;
                return DateTime.MinValue;
            }
        }

        public SPOUser CurrentUser
        {
            get
            {
                if (_currentUser == null && _web.IsObjectPropertyInstantiated("CurrentUser"))
                    _currentUser = new SPOUser(_web.CurrentUser);
                return _currentUser;
            }
        }

        public string CustomMasterUrl
        {
            get
            {
                if (_web.IsPropertyAvailable("CustomMasterUrl"))
                    return _web.CustomMasterUrl;
                return null;
            }
            set
            {
                _web.CustomMasterUrl = value;
                _hasChanges = true;
            }
        }

        public string Description
        {
            get
            {
                if (_web.IsPropertyAvailable("Description"))
                    return _web.Description;
                return null;
            }
            set
            {
                _web.Description = value;
                _hasChanges = true;
            }
        }

        public bool DocumentLibraryCalloutOfficeWebAppPreviewersDisabled
        {
            get
            {
                if (_web.IsPropertyAvailable("DocumentLibraryCalloutOfficeWebAppPreviewersDisabled"))
                    return _web.DocumentLibraryCalloutOfficeWebAppPreviewersDisabled;
                return false;
            }
        }

        //public BasePermissions EffectiveBasePermissions { get; private set; }
        public bool EnableMinimalDownload
        {
            get
            {
                if (_web.IsPropertyAvailable("EnableMinimalDownload"))
                    return _web.EnableMinimalDownload;
                return false;
            }
            set
            {
                _web.EnableMinimalDownload = value;
                _hasChanges = true;
            }
        }

        //public FeatureCollection Features { get; private set; }
        //public FieldCollection Fields { get; private set; }

        public bool HasUniqueRoleAssignments
        {
            get
            {
                if (_web.IsPropertyAvailable("HasUniqueRoleAssignments"))
                    return _web.HasUniqueRoleAssignments;
                return false;
            }
        }

        public Guid Id
        {
            get
            {
                if (_web.IsPropertyAvailable("Id"))
                    return _web.Id;
                return Guid.Empty;
            }
        }

        public uint Language
        {
            get
            {
                if (_web.IsPropertyAvailable("Language"))
                    return _web.Language;
                return 0;
            }
        }

        public DateTime LastItemModifiedDate
        {
            get
            {
                if (_web.IsPropertyAvailable("LastItemModifiedDate"))
                    return _web.LastItemModifiedDate;
                return DateTime.MinValue;
            }
        }

        //public ListTemplateCollection ListTemplates { get; private set; }
        public string MasterUrl
        {
            get
            {
                if (_web.IsPropertyAvailable("MasterUrl"))
                    return _web.MasterUrl;
                return null;
            }
            set
            {
                _web.MasterUrl = value;
                _hasChanges = true;
            }
        }


        public SPONavigation Navigation
        {
            get
            {
                if (_navigation == null && _web.IsObjectPropertyInstantiated("Navigation"))
                {
                    SPONavigation.LoadNavigation(SPOSiteContext.CurrentSiteContext.Context, _web.Navigation);
                    _navigation = new SPONavigation(_web.Navigation);
                }
                return _navigation;
            }
        }
        //public Navigation Navigation { get; private set; }
        public bool QuickLaunchEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("QuickLaunchEnabled"))
                    return _web.QuickLaunchEnabled;
                return false;
            }
            set
            {
                _web.QuickLaunchEnabled = value;
                _hasChanges = true;
            }
        }

        public RecycleBinItemCollection RecycleBin
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("RecycleBin"))
                    return _web.RecycleBin;
                return null;
            }
        }

        public bool RecycleBinEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("RecycleBinEnabled"))
                    return _web.RecycleBinEnabled;
                return false;
            }
        }

        public RegionalSettings RegionalSettings
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("RegionalSettings"))
                    return _web.RegionalSettings;
                return null;
            }
        }

        public RoleDefinitionCollection RoleDefinitions
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("RoleDefinitions"))
                    return _web.RoleDefinitions;
                return null;
            }
        }

        //public Folder RootFolder { get; private set; }
        public bool SaveSiteAsTemplateEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("SaveSiteAsTemplateEnabled"))
                    return _web.SaveSiteAsTemplateEnabled;
                return false;
            }
            set
            {
                _web.SaveSiteAsTemplateEnabled = value;
                _hasChanges = true;
            }
        }

        public string ServerRelativeUrl
        {
            get
            {
                if (_web.IsPropertyAvailable("ServerRelativeUrl"))
                    return _web.ServerRelativeUrl;
                return null;
            }
            set
            {
                _web.ServerRelativeUrl = value;
                _hasChanges = true;
            }
        }

        public bool ShowUrlStructureForCurrentUser
        {
            get
            {
                if (_web.IsPropertyAvailable("ShowUrlStructureForCurrentUser"))
                    return _web.ShowUrlStructureForCurrentUser;
                return false;
            }
        }

        public string AlternateCssUrl
        {
            get
            {
                if (_web.IsPropertyAvailable("AlternateCssUrl"))
                    return _web.AlternateCssUrl;
                return null;
            }
            set
            {
                _web.AlternateCssUrl = value;
                _hasChanges = true;
            }
        }
        public string SiteLogoUrl
        {
            get
            {
                if (_web.IsPropertyAvailable("SiteLogoUrl"))
                    return _web.SiteLogoUrl;
                return null;
            }
            set
            {
                _web.SiteLogoUrl = value;
                _hasChanges = true;
            }
        }
        //public GroupCollection SiteGroups { get; private set; }
        //public List SiteUserInfoList { get; private set; }
        //public UserCollection SiteUsers { get; private set; }
        public IEnumerable<int> SupportedUILanguageIds
        {
            get
            {
                if (_web.IsPropertyAvailable("SupportedUILanguageIds"))
                    return _web.SupportedUILanguageIds;
                return null;
            }
        }

        public bool SyndicationEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("SyndicationEnabled"))
                    return _web.SyndicationEnabled;
                return false;
            }
            set
            {
                _web.SyndicationEnabled = value;
                _hasChanges = true;
            }
        }

        public ThemeInfo ThemeInfo
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("ThemeInfo"))
                    return _web.ThemeInfo;
                return null;
            }
        }

        public string Title
        {
            get
            {
                if (_web.IsPropertyAvailable("Title"))
                    return _web.Title;
                return null;
            }
            set
            {
                _web.Title = value;
                _hasChanges = true;
            }
        }

        public bool TreeViewEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("TreeViewEnabled"))
                    return _web.TreeViewEnabled;
                return false;
            }
            set
            {
                _web.TreeViewEnabled = value;
                _hasChanges = true;
            }
        }

        public int UIVersion
        {
            get
            {
                if (_web.IsPropertyAvailable("UIVersion"))
                    return _web.UIVersion;
                return 0;
            }
            set
            {
                _web.UIVersion = value;
                _hasChanges = true;
            }
        }

        public bool UIVersionConfigurationEnabled
        {
            get
            {
                if (_web.IsPropertyAvailable("UIVersionConfigurationEnabled"))
                    return _web.UIVersionConfigurationEnabled;
                return false;
            }
            set
            {
                _web.UIVersionConfigurationEnabled = value;
                _hasChanges = true;
            }
        }

        public string Url
        {
            get
            {
                if (_web.IsPropertyAvailable("Url"))
                    return _web.Url;
                return null;
            }
        }

        public SPOUserCustomActionCollection UserCustomActions
        {
            get
            {
                if (_userCustomActions != null)
                    return _userCustomActions;
                if (_web.IsObjectPropertyInstantiated("UserCustomActions"))
                {
                    _userCustomActions = new SPOUserCustomActionCollection(_web.UserCustomActions);
                    return _userCustomActions;
                }
                return null;
            }
        }

        //public WebCollection Webs { get; private set; }
        public string WebTemplate
        {
            get
            {
                if (_web.IsPropertyAvailable("WebTemplate"))
                    return _web.WebTemplate;
                return null;
            }
        }

        public WorkflowAssociationCollection WorkflowAssociations
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("WorkflowAssociations"))
                    return _web.WorkflowAssociations;
                return null;
            }
        }

        public WorkflowTemplateCollection WorkflowTemplates
        {
            get
            {
                if (_web.IsObjectPropertyInstantiated("WorkflowTemplates"))
                    return _web.WorkflowTemplates;
                return null;
            }
        }
        

        public List<SPOWeb> GetSubwebsForCurrentUser(bool detailed)
        {
            return GetSubwebsForCurrentUser(detailed, false);
        }
        public List<SPOWeb> GetSubwebsForCurrentUser(bool detailed, bool includeAppWebs)
        {
            if (_detailedSubWebsForCurrentUser != null) return _detailedSubWebsForCurrentUser;
            if (!detailed && _subWebsForCurrentUser != null) return _subWebsForCurrentUser;
            
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var subWebs = _web.GetSubwebsForCurrentUser(new SubwebQuery());
            ctx.Load(subWebs);
            ctx.ExecuteQuery();
            var webs = new List<SPOWeb>();
            foreach (Web childWeb in subWebs)
            {
                SPOWeb.LoadWeb(ctx, childWeb, detailed);
                webs.Add(new SPOWeb(childWeb));
            }
            _subWebsForCurrentUser = webs;
            if (detailed)
                _detailedSubWebsForCurrentUser = webs;

            if (includeAppWebs)
                return webs;
            return webs.Where(w => w.WebTemplate != "APP").ToList();
        }


        public List<SPOWeb> GetWebs(bool detailed)
        {
            return GetWebs(detailed, false);
        }
        public List<SPOWeb> GetWebs(bool detailed, bool includeAppWebs)
        {
            if (_detailedSubWebs != null) return _detailedSubWebs;
            if (!detailed && _subWebs != null) return _subWebs;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var subWebs = _web.Webs;
            ctx.Load(subWebs);
            ctx.ExecuteQuery();
            var webs = new List<SPOWeb>();
            foreach (Web childWeb in subWebs)
            {
                SPOWeb.LoadWeb(ctx, childWeb, detailed);
                webs.Add(new SPOWeb(childWeb));
            }
            _subWebs = webs;
            if (detailed)
                _detailedSubWebs = webs;

            if (includeAppWebs)
                return webs;
            return webs.Where(w => w.WebTemplate != "APP").ToList();
        }

        private void IterateWebs(Web parentWeb, List<SPOWeb> webs, bool detailed)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var childWebs = parentWeb.Webs;
            ctx.Load(childWebs);
            ctx.ExecuteQuery();
            foreach (Web childWeb in childWebs)
            {
                SPOWeb.LoadWeb(ctx, childWeb, detailed);
                webs.Add(new SPOWeb(childWeb));
                IterateWebs(childWeb, webs, detailed);
            }
        }

        public List<SPOWeb> GetAllChildWebs(bool detailed)
        {
            return GetAllChildWebs(detailed, false);
        }
        public List<SPOWeb> GetAllChildWebs(bool detailed, bool includeAppWebs)
        {
            if (_detailedAllSubWebs != null) return _detailedAllSubWebs;
            if (!detailed && _allSubWebs != null) return _allSubWebs;

            var webs = new List<SPOWeb>();
            IterateWebs(_web, webs, detailed);

            _allSubWebs = webs;
            if (detailed)
                _detailedAllSubWebs = webs;

            if (includeAppWebs)
                return webs;
            return webs.Where(w => w.WebTemplate != "APP").ToList();
        }

         
        public List<SPOEventReceiverDefinition> GetEventReceivers()
        {
            if (_eventReceivers != null) return _eventReceivers;

            var eventReceivers = _web.EventReceivers;
            _web.Context.Load(eventReceivers);
            _web.Context.ExecuteQuery();
            _eventReceivers = new List<SPOEventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in eventReceivers)
            {
                _eventReceivers.Add(new SPOEventReceiverDefinition(erd, _web.EventReceivers));
            }
            return _eventReceivers;
        }

        public SPOEventReceiverDefinition AddEventReceiver(EventReceiverDefinitionCreationInformation erdci)
        {
            SPOEventReceiverDefinition erd = new SPOEventReceiverDefinition(_web.EventReceivers.Add(erdci), _web.EventReceivers);
            _web.EventReceivers.Context.ExecuteQuery();
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
            _web.BreakRoleInheritance(copyRoleAssignments, clearScopes);
            _web.Context.ExecuteQuery();
            Refresh();
        }

        public void ResetRoleInheritance()
        {
            _web.ResetRoleInheritance();
            _web.Context.ExecuteQuery();
            Refresh();
        }

        public void ApplyTheme(string colorPaletteUrl, string fontSchemeUrl, string backgroundImageUrl, bool shareGenerated)
        {
            if (string.IsNullOrEmpty(fontSchemeUrl))
                fontSchemeUrl = null;
            if (string.IsNullOrEmpty(backgroundImageUrl))
                backgroundImageUrl = null;
            _web.ApplyTheme(colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated);
            _web.Context.ExecuteQuery();
        }

 


        public void Refresh()
        {
            _web.RefreshLoad();
            _web.Context.ExecuteQuery();

            _detailedLists = null;
            _lists = null;
            _availableContentTypes = null;
            _contentTypes = null;
            _folders = null;
            _allProperties = null;
            _associatedMemberGroup = null;
            _associatedOwnerGroup = null;
            _associatedVisitorGroup = null;
            _currentUser = null;
            _subWebs = null;
            _subWebsForCurrentUser = null;
            _detailedSubWebsForCurrentUser = null;
            _detailedSubWebs = null;
            _allSubWebs = null;
            _detailedAllSubWebs = null;
            _eventReceivers = null;
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _web.Update();
                _web.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

        internal static void LoadWeb(CmdletContext ctx, Web web, bool detail)
        {
            ctx.Load(web);
            ctx.Load(web, w => w.HasUniqueRoleAssignments,
                w => w.EnableMinimalDownload);
            if (detail)
            {
                ctx.Load(web,
                    w => w.AllowDesignerForCurrentUser,
                    w => w.AllowMasterPageEditingForCurrentUser,
                    w => w.AllowRevertFromTemplateForCurrentUser,
                    w => w.AllProperties,
                    w => w.AlternateCssUrl,
                    w => w.AssociatedMemberGroup,
                    w => w.AssociatedOwnerGroup,
                    w => w.AssociatedVisitorGroup,
                    //w => w.AvailableContentTypes,
                    //w => w.AvailableFields,
                    //w => w.ContentTypes,
                    w => w.CurrentUser,
                    //w => w.EffectiveBasePermissions,
                    //w => w.EventReceivers,
                    //w => w.Features,
                    //w => w.Fields,
                    //w => w.FirstUniqueAncestorSecurableObject,
                    //w => w.Folders,
                    //w => w.Lists,
                    //w => w.ListTemplates,
                    w => w.Navigation,
                    w => w.RecycleBin,
                    w => w.RegionalSettings,
                    w => w.RoleAssignments,
                    w => w.RoleDefinitions,
                    //w => w.RootFolder,
                    w => w.SaveSiteAsTemplateEnabled,
                    w => w.ShowUrlStructureForCurrentUser,
                    //w => w.SiteGroups,
                    //w => w.SiteUserInfoList,
                    //w => w.SiteUsers,
                    w => w.SiteLogoUrl,
                    w => w.SupportedUILanguageIds,
                    w => w.ThemeInfo,
                    w => w.UserCustomActions,
                    w => w.WorkflowAssociations,
                    w => w.WorkflowTemplates);
            }
            ctx.ExecuteQuery();
        }


    }
}
