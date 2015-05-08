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
    [Cmdlet("New", "SPOTermGroup", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Creates a Term Group object.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> New-SPOTermGroup -TermStore 0 -Name \"My Group\"",
        Remarks = "This example creates a single term group.")]
    public class NewSPOTermGroup : BaseSPOCmdlet
    {
        public NewSPOTermGroup()
        {
            Id = Guid.NewGuid();
        }

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermStore to add the Term Group to. The type must be a valid TermStore object, name, GUID, or index location.")]
        public SPOTermStorePipeBind TermStore { get; set; }

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "The name of the term group to create.")]
        public string Name { get; set; }

        [Parameter(Mandatory = false, Position = 2, HelpMessage = "The Id of the term group to create.")]
        public Guid Id { get; set; }

        [Parameter(Mandatory = false, Position = 3, HelpMessage = "The description of the term group to create.")]
        public string Description { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            var termStore = new SPOTermStore(TermStore.Read());
            WriteObject(termStore.CreateGroup(Name, Id, Description));
        }
    }
}
