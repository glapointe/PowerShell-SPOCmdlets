using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public class SPOTermStore : SPOTaxonomyBase
    {
        private TermStore _termStore = null;
        private SPOTermGroup _systemGroup = null;
        private List<SPOTermGroup> _groups = null;
        private SPOTermSet _hashTagsTermSet = null;
        private SPOTermSet _orphanedTermsTermSet = null;
        private SPOTermSet _keywordsTermSet = null;

        public SPOTermStore(TermStore termStore)
        {
            _termStore = termStore;

            if (_termStore.ServerObjectIsNull == null)
            {
                _termStore.Context.Load(_termStore);
                _termStore.Context.Load(_termStore, ts => ts.ContentTypePublishingHub);
                _termStore.Context.ExecuteQuery();
            }

            if (_termStore.IsPropertyAvailable("DefaultLanguage"))
                SetProp("DefaultLanguage", _termStore.DefaultLanguage, false);
            if (_termStore.IsPropertyAvailable("WorkingLanguage"))
                SetProp("WorkingLanguage", _termStore.WorkingLanguage, false);

        }

        public TermStore TermStore { get { return _termStore; } }

        public Guid Id
        {
            get
            {
                if (_termStore.IsPropertyAvailable("Id"))
                    return _termStore.Id;
                return Guid.Empty;
            }
        }

        public string Name
        {
            get
            {
                if (_termStore.IsPropertyAvailable("Name"))
                    return _termStore.Name;
                return null;
            }
        }

        public string ContentTypePublishingHub
        {
            get
            {
                if (_termStore.IsPropertyAvailable("ContentTypePublishingHub"))
                    return _termStore.ContentTypePublishingHub;
                return null;
            }
        }

        public bool IsOnline
        {
            get
            {
                if (_termStore.IsPropertyAvailable("IsOnline"))
                    return _termStore.IsOnline;
                return false;
            }
        }

        public int DefaultLanguage
        {
            get
            {
                return GetPropV<int>("DefaultLanguage");
            }
            set
            {
                SetProp("DefaultLanguage", value, true);
            }
        }

        public int WorkingLanguage
        {
            get
            {
                return GetPropV<int>("WorkingLanguage");
            }
            set
            {
                SetProp("WorkingLanguage", value, true);
            }
        }

        public IEnumerable<int> Languages
        {
            get
            {
                if (_termStore.IsPropertyAvailable("Languages"))
                    return _termStore.Languages;
                return null;
            }
        }

        public SPOTermGroup GetSystemGroup()
        {
            if (_systemGroup != null) return _systemGroup;

            var systemGroup = _termStore.SystemGroup;
            if (!_termStore.IsObjectPropertyInstantiated("SystemGroup"))
            {
                _termStore.Context.Load(systemGroup);
                _termStore.Context.ExecuteQuery();
            }
            _systemGroup = new SPOTermGroup(systemGroup);
            return _systemGroup;
        }

        public List<SPOTermGroup> GetGroups()
        {
            if (_groups != null) return _groups;

            var groups = _termStore.Groups;
            _termStore.Context.Load(groups);
            _termStore.Context.ExecuteQuery();

            _groups = new List<SPOTermGroup>();
            foreach (TermGroup group in groups)
            {
                _groups.Add(new SPOTermGroup(group));
            }
            return _groups;
        }
        public SPOTermSet GetHashTagsTermSet()
        {
            if (_hashTagsTermSet != null) return _hashTagsTermSet;

            var hashTagsTermSet = _termStore.HashTagsTermSet;
            _termStore.Context.Load(hashTagsTermSet);
            _termStore.Context.ExecuteQuery();

            _hashTagsTermSet = new SPOTermSet(hashTagsTermSet);
            return _hashTagsTermSet;
        }

        public SPOTermSet GetOrphanedTermsTermSet()
        {
            if (_orphanedTermsTermSet != null) return _orphanedTermsTermSet;

            var orphanedTermsTermSet = _termStore.OrphanedTermsTermSet;
            _termStore.Context.Load(orphanedTermsTermSet);
            _termStore.Context.ExecuteQuery();

            _orphanedTermsTermSet = new SPOTermSet(orphanedTermsTermSet);
            return _orphanedTermsTermSet;
        }

        public SPOTermSet GetKeywordsTermSet()
        {
            if (_keywordsTermSet != null) return _keywordsTermSet;

            var keywordsTermSet = _termStore.KeywordsTermSet;
            _termStore.Context.Load(keywordsTermSet);
            _termStore.Context.ExecuteQuery();

            _keywordsTermSet = new SPOTermSet(keywordsTermSet);
            return _keywordsTermSet;
        }

        public SPOTermGroup CreateGroup(string name, Guid id, string description)
        {
            var ctx = _termStore.Context;
            TermGroup group = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(ctx);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    group = _termStore.Groups.GetByName(name);
                    ctx.Load(group);
                }
                using (scope.StartCatch())
                {
                }
            }
            ctx.ExecuteQuery();

            if (group == null || group.ServerObjectIsNull == null || group.ServerObjectIsNull.Value)
            {
                group = _termStore.CreateGroup(name, id);
                group.Description = description;
                ctx.ExecuteQuery();
                ctx.Load(group);
                ctx.ExecuteQuery();
                return new SPOTermGroup(group);
            }
            else
            {
                throw new Exception("The specified term group already exists.");
            }
        }

        public virtual void Refresh()
        {
            _termStore.RefreshLoad();
            _termStore.Context.ExecuteQuery();

            if (_termStore.IsPropertyAvailable("DefaultLanguage"))
                SetProp("DefaultLanguage", _termStore.DefaultLanguage, false);
            if (_termStore.IsPropertyAvailable("WorkingLanguage"))
                SetProp("WorkingLanguage", _termStore.WorkingLanguage, false);

            _systemGroup = null;
            _groups = null;
            _hashTagsTermSet = null;
            _orphanedTermsTermSet = null;
            _keywordsTermSet = null;
        }

        public virtual void Update()
        {
            if (HasChanges())
            {
                if (HasChange("DefaultLanguage"))
                    _termStore.DefaultLanguage = DefaultLanguage;
                if (HasChange("WorkingLanguage"))
                    _termStore.WorkingLanguage = WorkingLanguage;

                _termStore.Context.ExecuteQuery();
                Refresh();
            }
        }
    }
}
