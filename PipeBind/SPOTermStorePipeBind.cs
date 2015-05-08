using System;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOTermStorePipeBind : CmdletPipeBind<TermStore>
    {
        private Guid _termStoreId;
        private TermStore _termStore = null;

        public SPOTermStorePipeBind(TermStore termStore)
        {
            _termStoreId = termStore.Id;
        }
        public SPOTermStorePipeBind(SPOTermStore termStore)
        {
            _termStoreId = termStore.Id;
        }
        public SPOTermStorePipeBind(string name)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            _termStore = ts.TermStores.GetByName(name);
            ctx.Load(_termStore);
            ctx.Load(_termStore, ts1 => ts1.ContentTypePublishingHub);
            ctx.ExecuteQuery();
            _termStoreId = _termStore.Id;
        }
        public SPOTermStorePipeBind(Guid termStoreId)
        {
            _termStoreId = termStoreId;
        }
        public SPOTermStorePipeBind(int index)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            ctx.Load(ts.TermStores);
            ctx.ExecuteQuery();

            _termStore = ts.TermStores[0];
            ctx.Load(_termStore);
            ctx.Load(_termStore, ts1 => ts1.ContentTypePublishingHub);
            ctx.ExecuteQuery();
            _termStoreId = _termStore.Id;
        }

        protected override void Discover(TermStore instance)
        {
            _termStoreId = instance.Id;
        }

        public override TermStore Read()
        {
            if (_termStore != null)
                return _termStore;

            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = ts.TermStores.GetById(_termStoreId);
            ctx.Load(termStore);
            ctx.Load(termStore, ts1 => ts1.ContentTypePublishingHub);
            ctx.ExecuteQuery();
            return termStore;
        }
    }
}
