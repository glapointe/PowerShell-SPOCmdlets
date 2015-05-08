using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public class SPOTermGroup : SPOTaxonomyItem
    {
        private TermGroup _termGroup = null;

        public SPOTermGroup(TermGroup termGroup) : base(termGroup)
        {
            _termGroup = termGroup;

            if (_termGroup.IsPropertyAvailable("Description"))
                SetProp("Description", _termGroup.Description, false);
        }

        public TermGroup TermGroup { get { return _termGroup; } }

        public string Description
        {
            get
            {
                return GetProp<string>("Description");
            }
            set
            {
                SetProp("Description", value, true);
            }
        }

        public bool IsSiteCollectionGroup
        {
            get
            {
                if (_termGroup.IsPropertyAvailable("IsSiteCollectionGroup"))
                    return _termGroup.IsSiteCollectionGroup;
                return false;
            }
        }

        public bool IsSystemGroup
        {
            get
            {
                if (_termGroup.IsPropertyAvailable("IsSystemGroup"))
                    return _termGroup.IsSystemGroup;
                return false;
            }
        }

        private List<SPOTermSet> _termSets = null;
        public List<SPOTermSet> GetTermSets()
        {
            if (_termSets != null) return _termSets;

            var termSets = _termGroup.TermSets;
            _termGroup.Context.Load(termSets);
            _termGroup.Context.ExecuteQuery();

            _termSets = new List<SPOTermSet>();
            foreach (TermSet termSet in termSets)
            {
                _termSets.Add(new SPOTermSet(termSet));
            }
            return _termSets;
        }

        public void Delete()
        {
            _termGroup.DeleteObject();
            _termGroup.Context.ExecuteQuery();
        }

        public SPOTermSet CreateTermSet(string name, Guid id, int lcid, string contact, string description, string customSortOrder, bool isAvailableForTagging, string owner, bool isOpenForTermCreation)
        {
            var ctx = _termGroup.Context;
            TermSet termSet = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(ctx);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    termSet = _termGroup.TermSets.GetByName(name);
                    ctx.Load(termSet);
                }
                using (scope.StartCatch())
                {
                }
            }
            ctx.ExecuteQuery();

            if (termSet == null || termSet.ServerObjectIsNull == null || termSet.ServerObjectIsNull.Value)
            {
                if (lcid == 0)
                {
                    lcid = GetWorkingLanguage(_termGroup.TermStore);
                }
                if (string.IsNullOrEmpty(owner))
                {

                    var web = SPOSiteContext.CurrentSiteContext.Context.Web;
                    ctx.Load(web, w => w.CurrentUser);
                    ctx.ExecuteQuery();
                    owner = web.CurrentUser.LoginName;
                }

                termSet = _termGroup.CreateTermSet(name, id, lcid);

                termSet.Contact = contact;
                termSet.Description = description;
                termSet.CustomSortOrder = customSortOrder;
                termSet.IsAvailableForTagging = isAvailableForTagging;
                termSet.Owner = owner;
                termSet.IsOpenForTermCreation = isOpenForTermCreation;

                ctx.Load(termSet);
                ctx.ExecuteQuery();
                return new SPOTermSet(termSet);
            }
            else
            {
                throw new Exception("The specified term set already exists.");
            }
        }

        public override void Refresh()
        {
            _termGroup.RefreshLoad();
            _termGroup.Context.ExecuteQuery();
            base.Refresh();

            if (_termGroup.IsPropertyAvailable("Description"))
                SetProp("Description", _termGroup.Description, false);

            _termSets = null;
        }

        public override void Update()
        {
            if (HasChanges())
            {
                base.Update();

                if (HasChange("Description"))
                    _termGroup.Description = Description;
                _termGroup.Context.ExecuteQuery();
                Refresh();
            }
        }

    }
}
