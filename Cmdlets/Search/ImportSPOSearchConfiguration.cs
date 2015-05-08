using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Lapointe.SharePointOnline.PowerShell.Data;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client.Search.Administration;
using Microsoft.SharePoint.Client.Search.Portability;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Search
{
    [Cmdlet("Import", "SPOSearchConfiguration", SupportsShouldProcess = false)]
    [CmdletGroup("Search")]
    [CmdletDescription("Retrieves one or more Sites within the current Site Collection.", "Retrieves a single Site given the server relative URL, GUID, or existing SPOWeb object. If a Site identifier is not provided then all Sites are provided.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Import-SPOSearchConfiguration -InputObject \"c:\\search.xml\" -ConfigScope SPSite",
        Remarks = "This example imports a search configuration to the current Site Collection.")]
    public class ImportSPOSearchConfiguration : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Alias("Path"),
        Parameter(
            ValueFromPipeline = true,
            Position = 0,
            Mandatory = true,
            HelpMessage = "The path to the file containing the search settings to import. The value can be the path a file, an XmlDocument object, or an XML string.")]
        public XmlDocumentPipeBind InputObject { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "The level at which the search settings will be imported.")]
        public SearchObjectLevel ConfigScope { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            SearchObjectLevel configScope = ConfigScope;
            var searchConfigurationPortability = new SearchConfigurationPortability(ctx);
            var owner = new SearchObjectOwner(ctx, configScope);
            var searchConfigXml = InputObject.Read();
            searchConfigurationPortability.ImportSearchConfiguration(owner, searchConfigXml.OuterXml);
            ctx.ExecuteQuery();
        }
    }
}
