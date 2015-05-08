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
    [Cmdlet("Import", "SPOTaxonomy", SupportsShouldProcess = true, DefaultParameterSetName = "Tenant")]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Import the Managed Metadata Terms.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Import-SPOTaxonomy -Tenant -InputFile \"c:\\terms.xml\"",
        Remarks = "This example imports the terms from c:\\terms.xml to the Term Store associated with the tenant.")]
    [Example(Code = "PS C:\\> Import-SPOTaxonomy -ParentTermStore 0 -InputFile C:\\terms.xml -Verbose",
        Remarks = "This example imports the terms from c:\\terms.xml to the Term Store at index position 0.")]
    [Example(Code = "PS C:\\> Import-SPOTaxonomy -ParentGroup \"My Group\" -InputFile C:\\terms.xml -Verbose",
        Remarks = "This example imports the terms from c:\\terms.xml to the Group named My Group.")]
    [Example(Code = "PS C:\\> Import-SPOTaxonomy -ParentTermSet ([guid]\"6067b6a8-9866-40db-962d-84097f22619a\") -InputFile C:\\terms.xml -Verbose",
        Remarks = "This example imports the terms from c:\\terms.xml to the Term Set with ID \"6067b6a8-9866-40db-962d-84097f22619a\".")]
    [Example(Code = "PS C:\\> Import-SPOTaxonomy -ParentTerm ([guid]\"6067b6a8-9866-40db-962d-84097f22619a\") -InputFile C:\\terms.xml -Verbose",
        Remarks = "This example imports the terms from c:\\terms.xml to the Term with ID \"6067b6a8-9866-40db-962d-84097f22619a\".")]
    public sealed class ImportSPOTaxonomy : BaseSPOCmdlet
    {
        XmlDocument _xml = null;

        [Parameter(ParameterSetName = "Tenant",
            Mandatory = true,
            Position = 0,
            HelpMessage = "Import all terms for the tenant.")]
        public SwitchParameter Tenant { get; set; }

        [Parameter(ParameterSetName = "TermStore",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermStore object to import Groups into. The type must be a valid TermStore object, name, GUID, or index location.")]
        public SPOTermStorePipeBind ParentTermStore { get; set; }

        [Parameter(ParameterSetName = "Group",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Group object to import Term Sets into. The type must be a valid TermGroup object, SPOTermGroup object, name, or GUID.")]
        public SPOTermGroupPipeBind ParentGroup { get; set; }

        [Parameter(ParameterSetName = "TermSet",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The TermSet object to import Terms into. The type must be a valid TermSet object or GUID.")]
        public SPOTermSetPipeBind ParentTermSet { get; set; }

        [Parameter(ParameterSetName = "Term",
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 0,
            HelpMessage = "The Term object to import Terms into. The type must be a valid Term object or GUID.")]
        public SPOTermPipeBind ParentTerm { get; set; }


        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "The path to the file containing the terms to import or an XmlDocument object or XML string.",
            Position = 1)]
        [Alias("Xml", "Path")]
        public XmlDocumentPipeBind InputFile { get; set; }

        private void Validate()
        {
            _xml = InputFile.Read();
            string rootElement = _xml.DocumentElement.Name;
            bool isValidRoot = false;

            switch (ParameterSetName)
            {
                case "Tenant":
                    isValidRoot = rootElement == "TermStores" || rootElement == "TermStore";
                    break;
                case "TermStore":
                    isValidRoot = rootElement == "Groups" || rootElement == "Group";
                    break;
                case "Group":
                    isValidRoot = rootElement == "TermSets" || rootElement == "TermSet";
                    break;
                case "TermSet":
                    isValidRoot = rootElement == "Terms" || rootElement == "Term";
                    break;
                case "Term":
                    isValidRoot = rootElement == "Terms" || rootElement == "Term";
                    break;
            }
            if (!isValidRoot)
            {
                string msg = "The import file cannot be imported to the specified target location. The following details the allowed import targets:\r\n";
                msg += "\tTenant: Term Store\r\n";
                msg += "\tTerm Store: Group\r\n";
                msg += "\tGroup: Term Set\r\n";
                msg += "\tTerm Set: Term\r\n";
                msg += "\tTerm: Term";
                throw new Exception(msg);
            }
        }

        protected override void ProcessRecord()
        {
            if (!base.ShouldProcess("Import"))
                return;

            var ctx = base.Context;

            Validate();
            try
            {
                WriteVerbose(string.Format("Start Time: {0}", DateTime.Now.ToString()));

                ImportTermsHelper import = new ImportTermsHelper(this, ctx, _xml);

                switch (ParameterSetName)
                {
                    case "Tenant":
                        import.Import();
                        break;
                    case "TermStore":
                        import.Import(ParentTermStore.Read());
                        break;
                    case "Group":
                        import.Import(ParentGroup.Read());
                        break;
                    case "TermSet":
                        import.Import(ParentTermSet.Read());
                        break;
                    case "Term":
                        import.Import(ParentTerm.Read());
                        break;
                }
            }
            finally
            {
                WriteVerbose(string.Format("Finish Time: {0}", DateTime.Now.ToString()));
            }
        }


    }

}
