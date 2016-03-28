using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.EventReceivers;
using Lapointe.SharePointOnline.PowerShell.Data.Principals;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Sites
{
    public class SPOSite : SPOClientObject
    {
        private Site _site = null;
        private SPOWeb _rootWeb = null;
        private SPOWeb _rootWebDetailed = null;
        private SPOWeb _web = null;
        private SPOWeb _webDetailed = null;
        private bool _hasChanges = false;
        private SPOUserCustomActionCollection _userCustomActions = null;

        public SPOSite(Site site)
        {
            _site = site;
        }

        public Site Site { get { return _site; } }

        public bool AllowDesigner 
        { 
            get 
            {
                if (_site.IsPropertyAvailable("AllowDesigner"))
                    return _site.AllowDesigner;
                return false;
            }
            set
            {
                _site.AllowDesigner = value;
                _hasChanges = true;
            }
        }
        public bool AllowMasterPageEditing
        {
            get
            {
                if (_site.IsPropertyAvailable("AllowMasterPageEditing"))
                    return _site.AllowMasterPageEditing;
                return false;
            }
            set
            {
                _site.AllowMasterPageEditing = value;
                _hasChanges = true;
            }
        }
        public bool AllowRevertFromTemplate
        {
            get
            {
                if (_site.IsPropertyAvailable("AllowRevertFromTemplate"))
                    return _site.AllowRevertFromTemplate;
                return false;
            }
            set
            {
                _site.AllowRevertFromTemplate = value;
                _hasChanges = true;
            }
        }
        public bool AllowSelfServiceUpgrade
        {
            get
            {
                if (_site.IsPropertyAvailable("AllowSelfServiceUpgrade"))
                    return _site.AllowSelfServiceUpgrade;
                return false;
            }
            set
            {
                _site.AllowSelfServiceUpgrade = value;
                _hasChanges = true;
            }
        }
        public bool AllowSelfServiceUpgradeEvaluation
        {
            get
            {
                if (_site.IsPropertyAvailable("AllowSelfServiceUpgradeEvaluation"))
                    return _site.AllowSelfServiceUpgradeEvaluation;
                return false;
            }
            set
            {
                _site.AllowSelfServiceUpgradeEvaluation = value;
                _hasChanges = true;
            }
        }
        public bool CanUpgrade
        {
            get
            {
                if (_site.IsPropertyAvailable("CanUpgrade"))
                    return _site.CanUpgrade;
                return false;
            }
        }
        public int CompatibilityLevel
        {
            get
            {
                if (_site.IsPropertyAvailable("CompatibilityLevel"))
                    return _site.CompatibilityLevel;
                return 0;
            }
        }
        //public FeatureCollection Features { get; private set; }
        public Guid Id
        {
            get
            {
                if (_site.IsPropertyAvailable("Id"))
                    return _site.Id;
                return Guid.Empty;
            }
        }
        public string LockIssue
        {
            get
            {
                if (_site.IsPropertyAvailable("LockIssue"))
                    return _site.LockIssue;
                return null;
            }
        }
        public uint MaxItemsPerThrottledOperation
        {
            get
            {
                if (_site.IsPropertyAvailable("MaxItemsPerThrottledOperation"))
                    return _site.MaxItemsPerThrottledOperation;
                return 0;
            }
        }

        private SPOUser _owner = null;
        public SPOUser Owner
        {
            get
            {
                if (_owner == null && _site.IsObjectPropertyInstantiated("Owner"))
                    _owner = new SPOUser(_site.Owner);
                return _owner;
            }
        }
        public string PrimaryUri
        {
            get
            {
                if (_site.IsPropertyAvailable("PrimaryUri"))
                    return _site.PrimaryUri;
                return null;
            }
        }
        public bool ReadOnly
        {
            get
            {
                if (_site.IsPropertyAvailable("ReadOnly"))
                    return _site.ReadOnly;
                return false;
            }
        }
        //public RecycleBinItemCollection RecycleBin
        //{
        //    get
        //    {
        //        if (_site.IsObjectPropertyInstantiated("RecycleBin"))
        //            return _site.RecycleBin;
        //        return null;
        //    }
        //}
        public string ServerRelativeUrl
        {
            get
            {
                if (_site.IsPropertyAvailable("ServerRelativeUrl"))
                    return _site.ServerRelativeUrl;
                return null;
            }
        }
        public bool ShareByLinkEnabled
        {
            get
            {
                if (_site.IsPropertyAvailable("ShareByLinkEnabled"))
                    return _site.ShareByLinkEnabled;
                return false;
            }
        }
        public bool ShowUrlStructure
        {
            get
            {
                if (_site.IsPropertyAvailable("ShowUrlStructure"))
                    return _site.ShowUrlStructure;
                return false;
            }
            set
            {
                _site.ShowUrlStructure = value;
                _hasChanges = true;
            }
        }
        public bool UIVersionConfigurationEnabled
        {
            get
            {
                if (_site.IsPropertyAvailable("UIVersionConfigurationEnabled"))
                    return _site.UIVersionConfigurationEnabled;
                return false;
            }
            set
            {
                _site.UIVersionConfigurationEnabled = value;
                _hasChanges = true;
            }
        }
        public UpgradeInfo UpgradeInfo
        {
            get
            {
                if (_site.IsPropertyAvailable("UpgradeInfo"))
                    return _site.UpgradeInfo;
                return null;
            }
        }
        public DateTime UpgradeReminderDate
        {
            get
            {
                if (_site.IsPropertyAvailable("UpgradeReminderDate"))
                    return _site.UpgradeReminderDate;
                return DateTime.MinValue;
            }
        }
        public bool Upgrading
        {
            get
            {
                if (_site.IsPropertyAvailable("Upgrading"))
                    return _site.Upgrading;
                return false;
            }
        }
        public string Url
        {
            get
            {
                if (_site.IsPropertyAvailable("Url"))
                    return _site.Url;
                return null;
            }
        }
        public UsageInfo Usage
        {
            get
            {
                if (_site.IsPropertyAvailable("Usage"))
                    return _site.Usage;
                return null;
            }
        }
        public SPOUserCustomActionCollection UserCustomActions
        {
            get
            {
                if (_userCustomActions != null)
                    return _userCustomActions;
                if (_site.IsObjectPropertyInstantiated("UserCustomActions"))
                {
                    _userCustomActions = new SPOUserCustomActionCollection(_site.UserCustomActions);
                    return _userCustomActions;
                }
                return null;
            }
        }

        private List<SPOEventReceiverDefinition> _eventReceivers = null;
        public List<SPOEventReceiverDefinition> GetEventReceivers()
        {
            if (_eventReceivers != null) return _eventReceivers;

            var eventReceivers = _site.EventReceivers;
            _site.Context.Load(eventReceivers);
            _site.Context.ExecuteQuery();
            _eventReceivers = new List<SPOEventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in eventReceivers)
            {
                _eventReceivers.Add(new SPOEventReceiverDefinition(erd, _site.EventReceivers));
            }
            return _eventReceivers;
        }

        public SPOEventReceiverDefinition AddEventReceiver(EventReceiverDefinitionCreationInformation erdci)
        {
            SPOEventReceiverDefinition erd = new SPOEventReceiverDefinition(_site.EventReceivers.Add(erdci), _site.EventReceivers);
            _site.EventReceivers.Context.ExecuteQuery();
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

        public void Refresh()
        {
            _site.RefreshLoad();
            _site.Context.ExecuteQuery();
        }


        public void Update()
        {
            if (_hasChanges)
            {
                _site.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

        public SPOWeb GetWeb(bool detailed)
        {
            if (_webDetailed != null) return _webDetailed;
            if (_web != null && !detailed) return _web;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var web = ctx.Web;
            SPOWeb.LoadWeb(ctx, web, detailed);
            ctx.ExecuteQuery();

            var spoWeb = new SPOWeb(web);
            if (detailed)
                _webDetailed = spoWeb;
            else
                _web = spoWeb;

            return spoWeb;
        }


        public SPOWeb GetRootWeb(bool detailed)
        {
            if (_rootWebDetailed != null) return _rootWebDetailed;
            if (_rootWeb != null && !detailed) return _rootWeb;
            
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            var rootWeb = this.Site.RootWeb;
            SPOWeb.LoadWeb(ctx, rootWeb, detailed);
            ctx.ExecuteQuery();

            var spoWeb = new SPOWeb(rootWeb);
            if (detailed)
                _rootWebDetailed = spoWeb;
            else
                _rootWeb = spoWeb;

            return spoWeb;
        }

        internal static void LoadSite(CmdletContext ctx, Site site, bool detailed)
        {
            ctx.Load(site);
            if (detailed)
            {
                // Some new bug is preventing the RecycleBin property from being populated.
                ctx.Load(site,
                    s => s.CanUpgrade,
                    s => s.Owner,
                    //s => s.RecycleBin,
                    s => s.UpgradeInfo,
                    s => s.Usage,
                    s => s.UserCustomActions);
            }
            ctx.ExecuteQuery();
        }
    }
}
