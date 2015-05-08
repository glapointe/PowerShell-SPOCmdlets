using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data.Sites
{
    public class SPOUserCustomActionCollection : SPOClientObjectCollection<SPOUserCustomAction>
    {
        private UserCustomActionCollection _userCustomActionCollection;

        public SPOUserCustomActionCollection(UserCustomActionCollection userCustomActionCollection)
        {
            _userCustomActionCollection = userCustomActionCollection;
            foreach (var uca in userCustomActionCollection)
            {
                AddChild(new SPOUserCustomAction(uca));
            }
        }

        public UserCustomActionCollection UserCustomActionCollection { get { return _userCustomActionCollection; } }


        public SPOUserCustomAction Add()
        {
            var uca = new SPOUserCustomAction(_userCustomActionCollection.Add());
            AddChild(uca);
            return uca;
        }

        internal void Remove(SPOUserCustomAction uca)
        {
            Data.Remove(uca);
        }

        public void Clear()
        {
            _userCustomActionCollection.Clear();
            _userCustomActionCollection.Context.ExecuteQuery();
            Data.Clear();
        }

        public SPOUserCustomAction GetById(Guid id)
        {
            var uca = _userCustomActionCollection.GetById(id);
            _userCustomActionCollection.Context.ExecuteQuery();
            return new SPOUserCustomAction(uca);
        }
    }
}
