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
    [Cmdlet("Get", "SPOTermGroup", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Retrieves a Term Group object.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Get-SPOTermGroup -Identity \"My Group\"",
        Remarks = "This example retrieves a single term group.")]
    public class GetSPOTermGroup : BaseSPOCmdlet
    {
        [Alias("Group")]
        [Parameter(Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The taxonomy group to retrieve. The type must be a valid TermGroup object, SPOTermGroup object, name, or GUID.")]
        public SPOTermGroupPipeBind Identity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Identity != null)
            {
                var group = new SPOTermGroup(Identity.Read());
                WriteObject(group);
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
                        WriteObject(new SPOTermGroup(group));
                    }
                }
            }
        }

    }
}
