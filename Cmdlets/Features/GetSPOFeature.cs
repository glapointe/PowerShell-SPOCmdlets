using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Features
{
    [Cmdlet("Get", "SPOFeature", SupportsShouldProcess = false)]
    [CmdletGroup("Features")]
    [CmdletDescription("Retrieves the Feature definition IDs for the Site Collection or a specific Site.",
        "If the -Web parameter is not provided then the Site Collection scoped Feature definition IDs are returned; otherwise, the Site scoped Feature definition IDs for a given Site are returned.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOFeature",
        Remarks = "Retrieves all Site Collection scoped Feature definition IDs.")]
    [Example(Code = "PS C:\\> Get-SPOFeature -Web \"/\"",
        Remarks = "Retrieves all Site scoped Feature definition IDs for the root Site.")]
    public class GetSPOFeature : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the Site whose site-scoped Features will be retrieved.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false)]
        public Guid Identity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Web == null)
            {
                var features = ctx.Site.Features;
                ctx.Load(features);
                ctx.ExecuteQuery();
                foreach (Feature feature in features)
                {
                    if (Identity == null || Identity == Guid.Empty || Identity == feature.DefinitionId)
                        WriteObject(feature.DefinitionId);
                }
            }
            else
            {
                string surl = Web.Read();
                Web web = ctx.Site.OpenWeb(surl);
                var features = web.Features;
                ctx.Load(features);
                ctx.ExecuteQuery();

                foreach (Feature feature in features)
                {
                    if (Identity == null || Identity == Guid.Empty || Identity == feature.DefinitionId)
                        WriteObject(feature.DefinitionId);
                }
            }
        }

    }
}
