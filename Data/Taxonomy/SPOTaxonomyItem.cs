using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public abstract class SPOTaxonomyItem : SPOTaxonomyBase
    {
        private TaxonomyItem _item = null;

        public SPOTaxonomyItem(TaxonomyItem item)
        {
            _item = item;

            if (_item.IsPropertyAvailable("Name"))
                SetProp("Name", _item.Name, false);
        }

        public string Name
        {
            get
            {
                return GetProp<string>("Name");
            }
            set
            {
                SetProp("Name", value, true);
            }
        }

        public DateTime CreatedDate
        {
            get
            {
                if (_item.IsPropertyAvailable("CreatedDate"))
                    return _item.CreatedDate;
                return DateTime.MinValue;
            }
        }

        public Guid Id
        {
            get
            {
                if (_item.IsPropertyAvailable("Id"))
                    return _item.Id;
                return Guid.Empty;
            }
        }

        public DateTime LastModifiedDate
        {
            get
            {
                if (_item.IsPropertyAvailable("LastModifiedDate"))
                    return _item.LastModifiedDate;
                return DateTime.MinValue;
            }
        }

        private SPOTermStore _termStore = null;
        public SPOTermStore GetTermStore()
        {
            if (_termStore != null) return _termStore;

            var termStore = _item.TermStore;
            if (!_item.IsObjectPropertyInstantiated("TermStore"))
            {
                _item.Context.Load(termStore);
                _item.Context.ExecuteQuery();
            }
            _termStore = new SPOTermStore(termStore);
            return _termStore;
        }


        public virtual void Refresh()
        {
            if (_item.IsPropertyAvailable("Name"))
                SetProp("Name", _item.Name, false);
        }

        public virtual void Update()
        {
            if (HasChanges())
            {
                if (HasChange("Name"))
                    _item.Name = Name;
            }
        }

    }
}
