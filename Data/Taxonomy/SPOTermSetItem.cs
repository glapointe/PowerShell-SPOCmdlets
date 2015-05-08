using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public abstract class SPOTermSetItem : SPOTaxonomyItem
    {
        private TermSetItem _termSetItem = null;

        public SPOTermSetItem(TermSetItem termSetItem) : base(termSetItem)
        {
            _termSetItem = termSetItem;

            if (_termSetItem.IsPropertyAvailable("CustomSortOrder"))
                SetProp("CustomSortOrder", _termSetItem.CustomSortOrder, false);
            if (_termSetItem.IsPropertyAvailable("IsAvailableForTagging"))
                SetProp("IsAvailableForTagging", _termSetItem.IsAvailableForTagging, false);
            if (_termSetItem.IsPropertyAvailable("Owner"))
                SetProp("Owner", _termSetItem.Owner, false);
        }

        public string CustomSortOrder
        {
            get
            {
                return GetProp<string>("CustomSortOrder");
            }
            set
            {
                SetProp("CustomSortOrder", value, true);
            }
        }

        public bool IsAvailableForTagging
        {
            get
            {
                return GetPropV<bool>("IsAvailableForTagging");
            }
            set
            {
                SetProp("IsAvailableForTagging", value, true);
            }
        }

        public string Owner
        {
            get
            {
                return GetProp<string>("Owner");
            }
            set
            {
                SetProp("Owner", value, true);
            }
        }

        public IDictionary<string, string> CustomProperties
        {
            get
            {
                if (_termSetItem.IsPropertyAvailable("CustomProperties"))
                    return _termSetItem.CustomProperties;
                return null;
            }
        }

        public void SetCustomProperty(string name, string value)
        {
            _termSetItem.SetCustomProperty(name, value);
            _termSetItem.Context.ExecuteQuery();
        }

        public void DeleteCustomProperty(string name)
        {
            _termSetItem.DeleteCustomProperty(name);
            _termSetItem.Context.ExecuteQuery();
        }

        public void Delete()
        {
            _termSetItem.DeleteObject();
            _termSetItem.Context.ExecuteQuery();
        }


        private List<SPOTerm> _terms = null;
        public List<SPOTerm> GetTerms()
        {
            if (_terms != null) return _terms;

            var terms = _termSetItem.Terms;
            _termSetItem.Context.Load(terms);
            _termSetItem.Context.ExecuteQuery();

            _terms = new List<SPOTerm>();
            foreach (Term term in terms)
            {
                _terms.Add(new SPOTerm(term));
            }
            return _terms;
        }
        public SPOTerm CreateTerm(string name)
        {
            return CreateTerm(name, GetWorkingLanguage(_termSetItem.TermStore));
        }
        public SPOTerm CreateTerm(string name, int lcid)
        {
            return CreateTerm(name, lcid, Guid.NewGuid());
        }
        public SPOTerm CreateTerm(string name, int lcid, Guid newTermId)
        {
            Term term = _termSetItem.CreateTerm(name, lcid, newTermId);
            _termSetItem.Context.ExecuteQuery();
            _termSetItem.Context.Load(term);
            _termSetItem.Context.ExecuteQuery();
            _terms = null;
            return new SPOTerm(term);
        }

        public override void Refresh()
        {
            base.Refresh();

            if (_termSetItem.IsPropertyAvailable("CustomSortOrder"))
                SetProp("CustomSortOrder", _termSetItem.CustomSortOrder, false);
            if (_termSetItem.IsPropertyAvailable("IsAvailableForTagging"))
                SetProp("IsAvailableForTagging", _termSetItem.IsAvailableForTagging, false);
            if (_termSetItem.IsPropertyAvailable("Owner"))
                SetProp("Owner", _termSetItem.Owner, false);

            _terms = null;
        }

        public override void Update()
        {
            if (HasChanges())
            {
                base.Update();

                if (HasChange("CustomSortOrder"))
                    _termSetItem.CustomSortOrder = CustomSortOrder;

                if (HasChange("IsAvailableForTagging"))
                    _termSetItem.IsAvailableForTagging = IsAvailableForTagging;

                if (HasChange("Owner"))
                    _termSetItem.Owner = Owner;
            }
        }

    }
}
