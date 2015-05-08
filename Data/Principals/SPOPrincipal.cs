using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace Lapointe.SharePointOnline.PowerShell.Data.Principals
{
    public class SPOPrincipal : SPOClientObject
    {
        private Principal _principal = null;
        private bool _hasChanges = false;

        public SPOPrincipal(Principal principal)
        {
            _principal = principal;
        }

        public Principal Principal { get { return _principal; } }

        public int Id
        {
            get
            {
                if (_principal.IsPropertyAvailable("Id"))
                    return _principal.Id;
                return 0;
            }
        }

        public bool IsHiddenInUI
        {
            get
            {
                if (_principal.IsPropertyAvailable("IsHiddenInUI"))
                    return _principal.IsHiddenInUI;
                return false;
            }
        }

        public string LoginName
        {
            get
            {
                if (_principal.IsPropertyAvailable("LoginName"))
                    return _principal.LoginName;
                return null;
            }
        }

        public PrincipalType PrincipalType
        {
            get
            {
                if (_principal.IsPropertyAvailable("PrincipalType"))
                    return _principal.PrincipalType;
                return Microsoft.SharePoint.Client.Utilities.PrincipalType.None;
            }
        }

        public string Title
        {
            get
            {
                if (_principal.IsPropertyAvailable("Title"))
                    return _principal.Title;
                return null;
            }
            set
            {
                _principal.Title = value;
                _hasChanges = true;
            }
        }

        public virtual void Update()
        {
            if (_hasChanges)
            {
                _principal.Context.ExecuteQuery();
                _hasChanges = true;
            }
        }

    }
}
