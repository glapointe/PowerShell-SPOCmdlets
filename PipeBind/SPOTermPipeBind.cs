using System;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOTermPipeBind : CmdletPipeBind<Term>
    {
        private Guid _termId;

        public SPOTermPipeBind(Term term)
        {
            _termId = term.Id;
        }
        public SPOTermPipeBind(SPOTerm term)
        {
            _termId = term.Id;
        }
        public SPOTermPipeBind(Guid termId)
        {
            _termId = termId;
        }

        protected override void Discover(Term instance)
        {
            _termId = instance.Id;
        }

        public override Term Read()
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            var term = ts.GetTerm(_termId);
            ctx.Load(term);
            ctx.ExecuteQuery();
            return term;
        }
    }
}
