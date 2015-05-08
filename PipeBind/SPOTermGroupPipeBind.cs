using System;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOTermGroupPipeBind : CmdletPipeBind<TermGroup>
    {
        private string _name = null;
        private Guid _groupId = Guid.Empty;

        public SPOTermGroupPipeBind(SPOTermGroup group)
        {
            _groupId = group.Id;
        }

        public SPOTermGroupPipeBind(TermGroup group)
        {
            _groupId = group.Id;
        }

        public SPOTermGroupPipeBind(Guid groupId)
        {
            _groupId = groupId;
        }
        public SPOTermGroupPipeBind(string name)
        {
            _name = name;
        }

        protected override void Discover(TermGroup instance)
        {
            _groupId = instance.Id;
        }
        


        public override TermGroup Read()
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            var termStores = ts.TermStores;
            ctx.Load(termStores);
            ctx.ExecuteQuery();

            foreach (TermStore termStore in termStores)
            {
                TermGroup group = null;
                if (!string.IsNullOrEmpty(_name))
                    group = termStore.Groups.GetByName(_name);
                else
                    group = termStore.Groups.GetById(_groupId);
                ctx.Load(group);
                ctx.ExecuteQuery();
                if (group != null && !group.ServerObjectIsNull.Value)
                    return group;
            }
            return null;
        }
    }
}
