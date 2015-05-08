using System;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOTermSetPipeBind : CmdletPipeBind<TermSet>
    {
        private Guid _termSetId;

        public SPOTermSetPipeBind(TermSet termSet)
        {
            _termSetId = termSet.Id;
        }
        public SPOTermSetPipeBind(SPOTermSet termSet)
        {
            _termSetId = termSet.Id;
        }
        public SPOTermSetPipeBind(Guid termSetId)
        {
            _termSetId = termSetId;
        }

        protected override void Discover(TermSet instance)
        {
            _termSetId = instance.Id;
        }

        public override TermSet Read()
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            var termStores = ts.TermStores;
            ctx.Load(termStores);
            ctx.ExecuteQuery();

            foreach (TermStore termStore in termStores)
            {
                TermSet termSet = termStore.GetTermSet(_termSetId);
                ctx.Load(termSet);
                ctx.ExecuteQuery();
                if (termSet != null && !termSet.ServerObjectIsNull.Value)
                    return termSet;
            }
            return null;

        }
    }
}
