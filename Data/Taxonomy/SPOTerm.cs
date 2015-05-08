using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public class SPOTerm : SPOTermSetItem
    {
        private Term _term = null;
        private SPOTermSet _termSet = null;
        private List<SPOTermSet> _termSets = null;
        private SPOTerm _parentTerm = null;
        private SPOTerm _sourceTerm = null;
        private SPOTermSet _pinSourceTermSet = null;
        private List<SPOTerm> _reusedTerms = null;
        private List<SPOTermLabel> _labels = null;

        public SPOTerm(Term term) : base(term)
        {
            _term = term;
        }

        public Term Term { get { return _term; } }

        public string Description
        {
            get
            {
                if (_term.IsPropertyAvailable("Description"))
                    return _term.Description;
                return null;
            }
        }

        public bool IsDeprecated
        {
            get
            {
                if (_term.IsPropertyAvailable("IsDeprecated"))
                    return _term.IsDeprecated;
                return false;
            }
        }

        public bool IsKeyword
        {
            get
            {
                if (_term.IsPropertyAvailable("IsKeyword"))
                    return _term.IsKeyword;
                return false;
            }
        }

        public bool IsPinned
        {
            get
            {
                if (_term.IsPropertyAvailable("IsPinned"))
                    return _term.IsPinned;
                return false;
            }
        }

        public bool IsPinnedRoot
        {
            get
            {
                if (_term.IsPropertyAvailable("IsPinnedRoot"))
                    return _term.IsPinnedRoot;
                return false;
            }
        }

        public bool IsReused
        {
            get
            {
                if (_term.IsPropertyAvailable("IsReused"))
                    return _term.IsReused;
                return false;
            }
        }

        public bool IsRoot
        {
            get
            {
                if (_term.IsPropertyAvailable("IsRoot"))
                    return _term.IsRoot;
                return false;
            }
        }

        public bool IsSourceTerm
        {
            get
            {
                if (_term.IsPropertyAvailable("IsSourceTerm"))
                    return _term.IsSourceTerm;
                return false;
            }
        }


        public IDictionary<string, string> LocalCustomProperties
        {
            get
            {
                if (_term.IsPropertyAvailable("LocalCustomProperties"))
                    return _term.LocalCustomProperties;
                return null;
            }
        }

        public IEnumerable<Guid> MergedTermIds
        {
            get
            {
                if (_term.IsPropertyAvailable("MergedTermIds"))
                    return _term.MergedTermIds;
                return null;
            }
        }

        public string PathOfTerm
        {
            get
            {
                if (_term.IsPropertyAvailable("PathOfTerm"))
                    return _term.PathOfTerm;
                return null;
            }
        }

        public int TermsCount
        {
            get
            {
                if (_term.IsPropertyAvailable("TermsCount"))
                    return _term.TermsCount;
                return 0;
            }
        }


        public SPOTermSet GetTermSet()
        {
            if (_termSet != null) return _termSet;

            var termSet = _term.TermSet;
            _term.Context.Load(termSet);
            _term.Context.ExecuteQuery();

            _termSet = new SPOTermSet(termSet);
            return _termSet;
        }

        public List<SPOTermSet> GetTermSets()
        {
            if (_termSets != null) return _termSets;

            var termSets = _term.TermSets;
            _term.Context.Load(termSets);
            _term.Context.ExecuteQuery();

            _termSets = new List<SPOTermSet>();
            foreach (TermSet termSet in termSets)
            {
                _termSets.Add(new SPOTermSet(termSet));
            }
            return _termSets;
        }

        public SPOTerm GetParentTerm()
        {
            if (_parentTerm != null) return _parentTerm;

            var parentTerm = _term.Parent;
            _term.Context.Load(parentTerm);
            _term.Context.ExecuteQuery();

            _parentTerm = new SPOTerm(parentTerm);
            return _parentTerm;
        }

        public SPOTerm GetSourceTerm()
        {
            if (_sourceTerm != null) return _sourceTerm;

            var sourceTerm = _term.SourceTerm;
            _term.Context.Load(sourceTerm);
            _term.Context.ExecuteQuery();

            _sourceTerm = new SPOTerm(sourceTerm);
            return _sourceTerm;
        }

        public SPOTermSet GetPinSourceTermSet()
        {
            if (_pinSourceTermSet != null) return _pinSourceTermSet;

            var pinSourceTermSet = _term.PinSourceTermSet;
            _term.Context.Load(pinSourceTermSet);
            _term.Context.ExecuteQuery();

            _pinSourceTermSet = new SPOTermSet(pinSourceTermSet);
            return _pinSourceTermSet;
        }

        public List<SPOTerm> GetReusedTerms()
        {
            if (_reusedTerms != null) return _reusedTerms;

            var reusedTerms = _term.ReusedTerms;
            _term.Context.Load(reusedTerms);
            _term.Context.ExecuteQuery();

            _reusedTerms = new List<SPOTerm>();
            foreach (Term reusedTerm in reusedTerms)
            {
                _reusedTerms.Add(new SPOTerm(reusedTerm));
            }
            return _reusedTerms;
        }

        public List<SPOTermLabel> GetLabels()
        {
            if (_labels != null) return _labels;

            var labels = _term.Labels;
            _term.Context.Load(labels);
            _term.Context.ExecuteQuery();

            _labels = new List<SPOTermLabel>();
            foreach (Label label in labels)
            {
                _labels.Add(new SPOTermLabel(label, this));
            }
            return _labels;
        }

        public SPOTermLabel CreateLabel(string labelName, bool isDefault)
        {
            GetWorkingLanguage(_term.TermStore);
            return CreateLabel(labelName, GetWorkingLanguage(_term.TermStore), isDefault);
        }

        public SPOTermLabel CreateLabel(string labelName, int lcid, bool isDefault)
        {
            Label label = _term.CreateLabel(labelName, lcid, isDefault);
            _term.Context.ExecuteQuery();
            _labels = null;
            return new SPOTermLabel(label, this);
        }

        public void Deprecate(bool doDeprecate)
        {
            _term.Deprecate(doDeprecate);
            _term.Context.ExecuteQuery();
        }

        public string GetDescription(int lcid)
        {
            var desc = _term.GetDescription(lcid);
            _term.Context.ExecuteQuery();
            return desc.Value;
        }

        public void SetDescription(string description)
        {
            _term.SetDescription(description, GetWorkingLanguage(_term.TermStore));
            _term.Context.ExecuteQuery();
        }

        public void SetDescription(string description, int lcid)
        {
            _term.SetDescription(description, lcid);
            _term.Context.ExecuteQuery();
        }

        public void SetLocalCustomProperty(string name, string value)
        {
            _term.SetLocalCustomProperty(name, value);
            _term.Context.ExecuteQuery();
        }

        public void DeleteLocalCustomProperty(string name)
        {
            _term.DeleteLocalCustomProperty(name);
            _term.Context.ExecuteQuery();
        }

        //public LabelCollection Labels { get; }


        public override void Refresh()
        {
            _term.RefreshLoad();
            _term.Context.ExecuteQuery();
            base.Refresh();

            _reusedTerms = null;
            _pinSourceTermSet = null;
            _sourceTerm = null;
            _parentTerm = null;
            _termSets = null;
            _termSet = null;
            _labels = null;
        }

        public override void Update()
        {
            if (HasChanges())
            {
                base.Update();
                _term.Context.ExecuteQuery();
                Refresh();
            }
        }
    }
}
