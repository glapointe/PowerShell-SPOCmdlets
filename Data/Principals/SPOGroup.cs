using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace Lapointe.SharePointOnline.PowerShell.Data.Principals
{
    public class SPOGroup : SPOPrincipal
    {
        private Group _group = null;
        private bool _hasChanges = false;

        public SPOGroup(Group group) : base(group)
        {
            _group = group;
        }

        public Group Group { get { return _group; } }

        public bool AllowMembersEditMembership
        {
            get
            {
                if (_group.IsPropertyAvailable("AllowMembersEditMembership"))
                    return _group.AllowMembersEditMembership;
                return false;
            }
            set
            {
                _group.AllowMembersEditMembership = value;
                _hasChanges = true;
            }
        }

        public bool AllowRequestToJoinLeave
        {
            get
            {
                if (_group.IsPropertyAvailable("AllowRequestToJoinLeave"))
                    return _group.AllowRequestToJoinLeave;
                return false;
            }
            set
            {
                _group.AllowRequestToJoinLeave = value;
                _hasChanges = true;
            }
        }

        public bool AutoAcceptRequestToJoinLeave
        {
            get
            {
                if (_group.IsPropertyAvailable("AutoAcceptRequestToJoinLeave"))
                    return _group.AutoAcceptRequestToJoinLeave;
                return false;
            }
            set
            {
                _group.AutoAcceptRequestToJoinLeave = value;
                _hasChanges = true;
            }
        }

        public bool CanCurrentUserEditMembership
        {
            get
            {
                if (_group.IsPropertyAvailable("CanCurrentUserEditMembership"))
                    return _group.CanCurrentUserEditMembership;
                return false;
            }
        }

        public bool CanCurrentUserManageGroup
        {
            get
            {
                if (_group.IsPropertyAvailable("CanCurrentUserManageGroup"))
                    return _group.CanCurrentUserManageGroup;
                return false;
            }
        }

        public bool CanCurrentUserViewMembership
        {
            get
            {
                if (_group.IsPropertyAvailable("CanCurrentUserViewMembership"))
                    return _group.CanCurrentUserViewMembership;
                return false;
            }
        }

        public string Description
        {
            get
            {
                if (_group.IsPropertyAvailable("Description"))
                    return _group.Description;
                return null;
            }
            set
            {
                _group.Description = value;
                _hasChanges = true;
            }
        }

        public bool OnlyAllowMembersViewMembership
        {
            get
            {
                if (_group.IsPropertyAvailable("OnlyAllowMembersViewMembership"))
                    return _group.OnlyAllowMembersViewMembership;
                return false;
            }
            set
            {
                _group.OnlyAllowMembersViewMembership = value;
                _hasChanges = true;
            }
        }

        public SPOPrincipal Owner
        {
            get
            {
                SPOPrincipal owner = null;
                if (_group.IsObjectPropertyInstantiated("Owner"))
                {
                    if (_group.Owner is Group)
                        owner = new SPOGroup((Group)_group.Owner);
                    else if (_group.Owner is User)
                        owner = new SPOUser((User)_group.Owner);
                    else
                        owner = new SPOPrincipal(_group.Owner);
                }
                return owner;
            }
            set
            {
                if (value != null)
                    _group.Owner = value.Principal;
                else
                    _group.Owner = null;
                _hasChanges = true;
            }
        }

        public string OwnerTitle
        {
            get
            {
                if (_group.IsPropertyAvailable("OwnerTitle"))
                    return _group.OwnerTitle;
                return null;
            }
        }

        public string RequestToJoinLeaveEmailSetting
        {
            get
            {
                if (_group.IsPropertyAvailable("RequestToJoinLeaveEmailSetting"))
                    return _group.RequestToJoinLeaveEmailSetting;
                return null;
            }
            set
            {
                _group.RequestToJoinLeaveEmailSetting = value;
                _hasChanges = true;
            }
        }

        private List<SPOUser> _users = null;
        public List<SPOUser> GetUsers()
        {
            if (_users != null) return _users;

            UserCollection uc = _group.Users;
            if (!_group.IsObjectPropertyInstantiated("Users") || !_group.Users.AreItemsAvailable)
            {
                _group.Context.Load(uc);
                _group.Context.ExecuteQuery();
            }
            _users = new List<SPOUser>();
            foreach (User u in uc)
            {
                _users.Add(new SPOUser(u));
            }
            return _users;
        }

        public override void Update()
        {
            base.Update();
            if (_hasChanges)
            {
                _group.Update();
                _group.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

    }
}
