using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    [Cmdlet("New", "SPOTermSet", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Creates a Term Set object.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> New-SPOTermSet -Group \"My Group\" -Name \"My Term Set\"",
        Remarks = "This example creates a single term set.")]
    public class NewSPOTermSet : BaseSPOCmdlet
    {
        public NewSPOTermSet()
        {
            Id = Guid.NewGuid();
            LCID = 0;
        }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Term Group to add the Term Set to. The type must be a valid TermGroup object, SPOTermGroup object, name, or GUID.")]
        public SPOTermGroupPipeBind Group { get; set; }

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "The name of the term set to create.")]
        public string Name { get; set; }

        [Parameter(Mandatory = false, Position = 2, HelpMessage = "The Id of the term set to create.")]
        public Guid Id { get; set; }

        [Parameter(Mandatory = false, Position = 3, HelpMessage = "The description of the term set to create.")]
        public string Description { get; set; }

        [Parameter(Mandatory = false, Position = 4, HelpMessage = "The contact for the Term Set object.")]
        public string Contact { get; set; }

        [Parameter(Mandatory = false, Position = 5, HelpMessage = "The custom sort order of the Term Set object.")]
        public string CustomSortOrder { get; set; }

        [Parameter(Mandatory = false, Position = 6, HelpMessage = "If specified, the Term Set object should appear in the tagging user interface.")]
        public SwitchParameter IsAvailableForTagging { get; set; }

        [Parameter(Mandatory = false, Position = 7, HelpMessage = "The user who is responsible for the current TermSet.")]
        public string Owner { get; set; }

        [Parameter(Mandatory = false, Position = 8, HelpMessage = "If specified, the Term Set can be updated by all users; otherwise, just by the specified constributors for the Group object.")]
        public SwitchParameter IsOpenForTermCreation { get; set; }

        [Alias("Locale", "Language")]
        [Parameter(Mandatory = false, Position = 9, HelpMessage = "The language code identifier (LCID).")]
        public int LCID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            var group = new SPOTermGroup(Group.Read());
            WriteObject(group.CreateTermSet(Name, Id, LCID, Contact, Description, CustomSortOrder, IsAvailableForTagging, Owner, IsOpenForTermCreation));
        }

        private int GetWorkingLanguage(TaxonomySession ts)
        {
            var ctx = base.Context;
            ctx.Load(ts.TermStores);
            ctx.ExecuteQuery();

            TermStore termStore = ts.TermStores[0];
            ctx.Load(termStore, t => t.WorkingLanguage);
            ctx.ExecuteQuery();
            return termStore.WorkingLanguage;
        }

    }
}
