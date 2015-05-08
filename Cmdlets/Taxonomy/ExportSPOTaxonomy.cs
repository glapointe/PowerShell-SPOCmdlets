using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Xml;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    [Cmdlet("Export", "SPOTaxonomy", SupportsShouldProcess = false, DefaultParameterSetName = "Tenant")]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Export the Managed Metadata Terms.")]
    [RelatedCmdlets(typeof(ImportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Export-SPOTaxonomy -Tenant -OutputFile \"c:\\terms.xml\" -Verbose",
        Remarks = "This example exports the terms for all term stores associated with the tenant and saves to c:\\terms.xml.")]
    [Example(Code = "PS C:\\> Export-SPOTaxonomy -TermStore 0 -OutputFile C:\\terms.xml -Verbose",
        Remarks = "This example exports the term store at index position zero c:\\terms.xml.")]
    [Example(Code = "PS C:\\> Export-SPOTaxonomy -Group \"My Group\" -OutputFile C:\\terms.xml -Verbose",
        Remarks = "This example exports a single group and saves to c:\\terms.xml.")]
    [Example(Code = "PS C:\\> Export-SPOTaxonomy -TermSet ([guid]\"6067b6a8-9866-40db-962d-84097f22619a\") -OutputFile C:\\terms.xml -Verbose",
        Remarks = "This example exports a single term set and saves to c:\\terms.xml.")]
    [Example(Code = "PS C:\\> Export-SPOTaxonomy -Term ([guid]\"6067b6a8-9866-40db-962d-84097f22619a\") -OutputFile C:\\terms.xml -Verbose",
        Remarks = "This example exports a single term and saves to c:\\terms.xml.")]
    public class ExportSPOTaxonomy : BaseSPOCmdlet
    {
        [Parameter(ParameterSetName = "Tenant",
            Mandatory = true,
            Position = 0,
            HelpMessage = "Export all terms for the tenant.")]
        public SwitchParameter Tenant { get; set; }

        [Parameter(ParameterSetName = "TermStore",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermStore object containing the terms to export. The type must be a valid TermStore object, name, GUID, or index location.")]
        public SPOTermStorePipeBind TermStore { get; set; }

        [Parameter(ParameterSetName = "Group",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Group object containing the terms to export. The type must be a valid TermGroup object, SPOTermGroup object, name, or GUID.")]
        public SPOTermGroupPipeBind Group { get; set; }

        [Parameter(ParameterSetName = "TermSet",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermSet object containing the terms to export. The type must be a valid TermSet object, SPOTermSet object, or GUID.")]
        public SPOTermSetPipeBind TermSet { get; set; }

        [Parameter(ParameterSetName = "Term",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Term object containing the terms to export. The type must be a valid Term object or GUID.")]
        public SPOTermPipeBind Term { get; set; }


        [Parameter(Mandatory = false,
            HelpMessage = "The path to the file to save the terms to.",
            Position = 1)]
        [ValidateDirectoryExistsAndValidFileName]
        [Alias("Path")]
        public string OutputFile { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            ExportTermsHelper export = new ExportTermsHelper(ctx);
            XmlDocument xml = null;

            switch (ParameterSetName)
            {
                case "Tenant":
                    xml = export.Export();
                    break;
                case "TermStore":
                    xml = export.Export(TermStore.Read());
                    break;
                case "Group":
                    xml = export.Export(Group.Read());
                    break;
                case "TermSet":
                    xml = export.Export(TermSet.Read());
                    break;
                case "Term":
                    xml = export.Export(Term.Read());
                    break;
            }
            if (xml == null)
                return;

            if (!string.IsNullOrEmpty(OutputFile))
                xml.Save(OutputFile);
            else
                WriteObject(xml);
        }
    }
}
