using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

namespace Lapointe.SharePointOnline.PowerShell.Data.Principals
{
    public class SPOUser : SPOPrincipal
    {
        private User _user = null;
        private bool _hasChanges = false;

        public SPOUser(User user) : base(user)
        {
            _user = user;
        }

        public User User { get { return _user; } }

        public string Email
        {
            get
            {
                if (_user.IsPropertyAvailable("Email"))
                    return _user.Email;
                return null;
            }
            set
            {
                _user.Email = value;
                _hasChanges = true;
            }
        }

        public bool IsSiteAdmin
        {
            get
            {
                if (_user.IsPropertyAvailable("IsSiteAdmin"))
                    return _user.IsSiteAdmin;
                return false;
            }
            set
            {
                _user.IsSiteAdmin = value;
                _hasChanges = true;
            }
        }

        public UserIdInfo UserId
        {
            get
            {
                if (_user.IsPropertyAvailable("UserId"))
                    return _user.UserId;
                return null;
            }
        }

        public override void Update()
        {
            base.Update();
            if (_hasChanges)
            {
                _user.Update();
                _user.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }


        private List<SPOGroup> _groups = null;
        public List<SPOGroup> GetGroups()
        {
            if (_groups != null) return _groups;

            GroupCollection gc = _user.Groups;
            if (!_user.IsObjectPropertyInstantiated("Groups") || !_user.Groups.AreItemsAvailable)
            {
                _user.Context.Load(gc);
                _user.Context.ExecuteQuery();
            }
            _groups = new List<SPOGroup>();
            foreach (Group g in gc)
            {
                _groups.Add(new SPOGroup(g));
            }
            return _groups;
        }

    }
}
