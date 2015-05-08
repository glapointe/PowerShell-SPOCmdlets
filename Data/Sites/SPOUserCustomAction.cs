using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data.Sites
{
    public class SPOUserCustomAction : SPOClientObject
    {
        private UserCustomAction _userCustomAction = null;
        private bool _hasChanges = false;

        public SPOUserCustomAction(UserCustomAction userCustomAction)
        {
            _userCustomAction = userCustomAction;
        }

        public UserCustomAction UserCustomAction { get { return _userCustomAction; } }
        public string CommandUIExtension
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("CommandUIExtension"))
                    return _userCustomAction.CommandUIExtension;
                return null;
            }
            set
            {
                _userCustomAction.CommandUIExtension = value;
                _hasChanges = true;
            }
        }
        public string Description
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Description"))
                    return _userCustomAction.Description;
                return null;
            }
            set
            {
                _userCustomAction.Description = value;
                _hasChanges = true;
            }
        }
        public string Group
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Group"))
                    return _userCustomAction.Group;
                return null;
            }
            set
            {
                _userCustomAction.Group = value;
                _hasChanges = true;
            }
        }
        public Guid Id
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Id"))
                    return _userCustomAction.Id;
                return Guid.Empty;
            }
        }
        public string ImageUrl
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("ImageUrl"))
                    return _userCustomAction.ImageUrl;
                return null;
            }
            set
            {
                _userCustomAction.ImageUrl = value;
                _hasChanges = true;
            }
        }
        public string Location
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Location"))
                    return _userCustomAction.Location;
                return null;
            }
            set
            {
                _userCustomAction.Location = value;
                _hasChanges = true;
            }
        }
        public string Name
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Name"))
                    return _userCustomAction.Name;
                return null;
            }
            set
            {
                _userCustomAction.Name = value;
                _hasChanges = true;
            }
        }
        public string RegistrationId
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("RegistrationId"))
                    return _userCustomAction.RegistrationId;
                return null;
            }
            set
            {
                _userCustomAction.RegistrationId = value;
                _hasChanges = true;
            }
        }
        public UserCustomActionRegistrationType RegistrationType
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("RegistrationType"))
                    return _userCustomAction.RegistrationType;
                return UserCustomActionRegistrationType.None;
            }
            set
            {
                _userCustomAction.RegistrationType = value;
                _hasChanges = true;
            }
        }
        public BasePermissions Rights
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Rights"))
                    return _userCustomAction.Rights;
                return null;
            }
            set
            {
                _userCustomAction.Rights = value;
                _hasChanges = true;
            }
        }
        public UserCustomActionScope Scope
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Scope"))
                    return _userCustomAction.Scope;
                return UserCustomActionScope.Unknown;
            }
        }
        public string ScriptBlock
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("ScriptBlock"))
                    return _userCustomAction.ScriptBlock;
                return null;
            }
            set
            {
                _userCustomAction.ScriptBlock = value;
                _hasChanges = true;
            }
        }
        public string ScriptSrc
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("ScriptSrc"))
                    return _userCustomAction.ScriptSrc;
                return null;
            }
            set
            {
                _userCustomAction.ScriptSrc = value;
                _hasChanges = true;
            }
        }
        public int Sequence
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Sequence"))
                    return _userCustomAction.Sequence;
                return 0;
            }
            set
            {
                _userCustomAction.Sequence = value;
                _hasChanges = true;
            }
        }
        public string Title
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Title"))
                    return _userCustomAction.Title;
                return null;
            }
            set
            {
                _userCustomAction.Title = value;
                _hasChanges = true;
            }
        }
        public string Url
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("Url"))
                    return _userCustomAction.Url;
                return null;
            }
            set
            {
                _userCustomAction.Url = value;
                _hasChanges = true;
            }
        }
        public string VersionOfUserCustomAction
        {
            get
            {
                if (_userCustomAction.IsPropertyAvailable("VersionOfUserCustomAction"))
                    return _userCustomAction.VersionOfUserCustomAction;
                return null;
            }
        }

        public void Delete()
        {
            _userCustomAction.DeleteObject();
            _userCustomAction.Context.ExecuteQuery();
            if (ParentCollection != null && ParentCollection is SPOUserCustomActionCollection)
                ((SPOUserCustomActionCollection)ParentCollection).Remove(this);
        }
        public void Update()
        {
            if (_hasChanges)
            {
                _userCustomAction.Update();
                _userCustomAction.Context.Load(_userCustomAction);
                _userCustomAction.Context.ExecuteQuery();

                _hasChanges = false;
            }
        }

    }
}
