using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    [Cmdlet("Get", "SPOTermStore", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Retrieves a Term Store object.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Get-SPOTermStore -Identity \"My Group\"",
        Remarks = "This example retrieves a single term group.")]
    public class GetSPOTermStore : BaseSPOCmdlet
    {
        [Alias("TermStore")]
        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The taxonomy term store to retrieve. The type must be a valid TermStore object, name, GUID, or index location.")]
        public SPOTermStorePipeBind Identity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Identity != null)
            {
                var termStore = new SPOTermStore(Identity.Read());
                WriteObject(termStore);
            }
            else
            {
                TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
                var termStores = ts.TermStores;
                ctx.Load(termStores);
                ctx.ExecuteQuery();

                foreach (TermStore termStore in termStores)
                {
                    ctx.Load(termStore, ts1 => ts1.ContentTypePublishingHub);
                    ctx.ExecuteQuery();
                    WriteObject(new SPOTermStore(termStore));
                }
            }
        }

    }
}
