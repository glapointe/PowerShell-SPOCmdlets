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
using Lapointe.SharePointOnline.PowerShell.Attributes;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Search
{
    [Cmdlet("Export", "SPOSearchConfiguration", SupportsShouldProcess = false)]
    [CmdletGroup("Search")]
    [CmdletDescription("Exports the search configuration.", "Exports the search configuration from the current Tenant or Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(ImportSPOSearchConfiguration))]
    [Example(Code = "PS C:\\> Export-SPOSearchConfiguration -Path \"c:\\search.xml\" -ConfigScope SPSite",
        Remarks = "This example exports the search configuration for the current Site Collection.")]
    public class ExportSPOSearchConfiguration : BaseSPOCmdlet
    {
        [ValidateDirectoryExistsAndValidFileName,
        Parameter(
            Position = 0,
            Mandatory = true,
            HelpMessage = "The file path where the exported search settings will be saved.")]
        public string Path { get; set; }

        [Parameter(
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
            var result = searchConfigurationPortability.ExportSearchConfiguration(owner);
            ctx.ExecuteQuery();
            System.IO.File.WriteAllText(Path, result.Value);
        }
    }
}
