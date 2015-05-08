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
    [Cmdlet("Get", "SPOTermSet", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Retrieves a Term Set object.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Get-SPOTermSet -Identity ([guid]\"6067b6a8-9866-40db-962d-84097f22619a\")",
        Remarks = "This example retrieves a single term set.")]
    public class GetSPOTermSet : BaseSPOCmdlet
    {
        [Alias("TermSet")]
        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermSet to retrieve. The type must be a valid TermSet object, SPOTermSet object, or GUID.")]
        public SPOTermSetPipeBind Identity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Identity != null)
            {
                var termSet = new SPOTermSet(Identity.Read());
                WriteObject(termSet);
            }
            else
            {
                TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
                var termStores = ts.TermStores;
                ctx.Load(termStores);
                ctx.ExecuteQuery();

                foreach (TermStore termStore in termStores)
                {
                    var groups = termStore.Groups;
                    ctx.Load(groups);
                    ctx.ExecuteQuery();
                    foreach (var group in groups)
                    {
                        var termSets = group.TermSets;
                        ctx.Load(termSets);
                        ctx.ExecuteQuery();
                        foreach (var termSet in termSets)
                        {
                            WriteObject(new SPOTermSet(termSet));
                        }
                    }
                }
            }
        }
    }
}
