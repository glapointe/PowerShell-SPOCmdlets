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
    [Cmdlet("Disable", "SPOFeature", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [CmdletGroup("Features")]
    [CmdletDescription("Disables a Feature within the current Site Collection or a specific Site.",
        "If the -Web parameter is not provided then a Site Collection scoped Feature will be disabled; otherwise, a Site scoped Feature for a given Site will be disabled.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(EnableSPOFeature))]
    [Example(Code = "PS C:\\> Disable-SPOFeature -Identity \"{7AF6EA6F-2079-4310-B964-210A87F0D5A1}\"",
        Remarks = "Disables a Site Collection scoped Feature.")]
    [Example(Code = "PS C:\\> Disable-SPOFeature -Web \"/\" -Identity \"{7AF6EA6F-2079-4310-B964-210A87F0D5A1}\"",
        Remarks = "Disables a Site scoped Feature on the root Site.")]
    public class DisableSPOFeature : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the Site where the specified Feature will be deactivated.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true)]
        public Guid Identity { get; set; }

        [Parameter(Position = 2, Mandatory = false)]
        public SwitchParameter Force { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (!base.ShouldProcess(Identity.ToString(), "Remove Content Type."))
                return;

            if (Web == null)
            {
                var features = ctx.Site.Features;
                ctx.Load(features);
                ctx.ExecuteQuery();
                features.Remove(Identity, Force);
                ctx.ExecuteQuery();
            }
            else
            {
                string surl = Web.Read();
                Web web = ctx.Site.OpenWeb(surl);
                var features = web.Features;
                ctx.Load(features);
                ctx.ExecuteQuery();
                features.Remove(Identity, Force);
                ctx.ExecuteQuery();
            }
        }

    }
}
