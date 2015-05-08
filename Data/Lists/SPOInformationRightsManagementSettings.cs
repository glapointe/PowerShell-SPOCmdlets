using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOInformationRightsManagementSettings : SPOClientObject
    {
        protected InformationRightsManagementSettings _irmSettings = null;
        protected ClientObject _parent = null;
        private bool _hasChanges = false;

        public SPOInformationRightsManagementSettings(InformationRightsManagementSettings irmSettings, ClientObject parent)
        {
            _irmSettings = irmSettings;
            _parent = parent;
        }

        public InformationRightsManagementSettings InformationRightsManagementSettings { get { return _irmSettings; } }

        public void Reset()
        {
            _irmSettings.Reset();
            _parent.Context.ExecuteQuery();
        }

        public void Update()
        {
            if (!_hasChanges) return;

            _irmSettings.Update();
            _parent.Context.ExecuteQuery();
        }

        public bool AllowPrint
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("AllowPrint"))
                    return _irmSettings.AllowPrint;
                return false;
            }
            set
            {
                _irmSettings.AllowPrint = value;
                _hasChanges = true;
            }
        }

        public bool AllowScript
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("AllowScript"))
                    return _irmSettings.AllowScript;
                return false;
            }
            set
            {
                _irmSettings.AllowScript = value;
                _hasChanges = true;
            }
        }

        public bool AllowWriteCopy
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("AllowWriteCopy"))
                    return _irmSettings.AllowWriteCopy;
                return false;
            }
            set
            {
                _irmSettings.AllowWriteCopy = value;
                _hasChanges = true;
            }
        }

        public bool DisableDocumentBrowserView
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("DisableDocumentBrowserView"))
                    return _irmSettings.DisableDocumentBrowserView;
                return false;
            }
            set
            {
                _irmSettings.DisableDocumentBrowserView = value;
                _hasChanges = true;
            }
        }

        public int DocumentAccessExpireDays
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("DocumentAccessExpireDays"))
                    return _irmSettings.DocumentAccessExpireDays;
                return 0;
            }
            set
            {
                _irmSettings.DocumentAccessExpireDays = value;
                _hasChanges = true;
            }
        }

        public DateTime DocumentLibraryProtectionExpireDate
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("DocumentLibraryProtectionExpireDate"))
                    return _irmSettings.DocumentLibraryProtectionExpireDate;
                return DateTime.MinValue;
            }
            set
            {
                _irmSettings.DocumentLibraryProtectionExpireDate = value;
                _hasChanges = true;
            }
        }

        public bool EnableDocumentAccessExpire
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("EnableDocumentAccessExpire"))
                    return _irmSettings.EnableDocumentAccessExpire;
                return false;
            }
            set
            {
                _irmSettings.EnableDocumentAccessExpire = value;
                _hasChanges = true;
            }
        }

        public bool EnableDocumentBrowserPublishingView
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("EnableDocumentBrowserPublishingView"))
                    return _irmSettings.EnableDocumentBrowserPublishingView;
                return false;
            }
            set
            {
                _irmSettings.EnableDocumentBrowserPublishingView = value;
                _hasChanges = true;
            }
        }

        public bool EnableGroupProtection
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("EnableGroupProtection"))
                    return _irmSettings.EnableGroupProtection;
                return false;
            }
            set
            {
                _irmSettings.EnableGroupProtection = value;
                _hasChanges = true;
            }
        }

        public bool EnableLicenseCacheExpire
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("EnableLicenseCacheExpire"))
                    return _irmSettings.EnableLicenseCacheExpire;
                return false;
            }
            set
            {
                _irmSettings.EnableLicenseCacheExpire = value;
                _hasChanges = true;
            }
        }

        public string GroupName
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("GroupName"))
                    return _irmSettings.GroupName;
                return null;
            }
            set
            {
                _irmSettings.GroupName = value;
                _hasChanges = true;
            }
        }

        public int LicenseCacheExpireDays
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("LicenseCacheExpireDays"))
                    return _irmSettings.LicenseCacheExpireDays;
                return 0;
            }
            set
            {
                _irmSettings.LicenseCacheExpireDays = value;
                _hasChanges = true;
            }
        }

        public string PolicyDescription
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("PolicyDescription"))
                    return _irmSettings.PolicyDescription;
                return null;
            }
            set
            {
                _irmSettings.PolicyDescription = value;
                _hasChanges = true;
            }
        }

        public string PolicyTitle
        {
            get
            {
                if (_irmSettings.IsPropertyAvailable("PolicyTitle"))
                    return _irmSettings.PolicyTitle;
                return null;
            }
            set
            {
                _irmSettings.PolicyTitle = value;
                _hasChanges = true;
            }
        }
    }
}
