using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Publishing;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Solutions
{
    [Cmdlet("Install", "SPOSolution", SupportsShouldProcess = false)]
    [CmdletGroup("Solutions")]
    [CmdletDescription("Add a WSP to the Solution Gallery of a Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Install-SPOSolution -Identity \"customtemplate.wsp\"",
        Remarks = "This example activates a custom Sandbox Solution in the Solution Gallery of the current Site Collection. The WSP will be copied to customtemplate.wsp-v1.0 and this copied version will be activated.")]
    public class InstallSPOSolution : BaseSPOCmdlet
    {
        public InstallSPOSolution() { MajorVersion = 1; }

        [ValidateNotNullOrEmpty,
        Alias("PackageName", "Name"),
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "Specify the name of the solution package (WSP) to activate.")]
        public string Identity { get; set; }

        [Parameter(Position = 2, HelpMessage = "Sets the major version of the WSP. The default value is 1.")]
        public int MajorVersion { get; set; }

        [Parameter(Position = 3, HelpMessage = "Sets the minor version of the WSP. The default value is 0.")]
        public int MinorVersion { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            ctx.Load(ctx.Site, s => s.ServerRelativeUrl);
            ctx.ExecuteQuery();

            DesignPackageInfo dpi = new DesignPackageInfo();
            dpi.PackageGuid = Guid.Empty;
            dpi.MajorVersion = MajorVersion;
            dpi.MinorVersion = MinorVersion;
            dpi.PackageName = Identity;

            string url = ctx.Site.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/solutions/" + Identity;
            DesignPackage.Install(ctx, ctx.Site, dpi, url);
            ctx.ExecuteQuery();


        }
    }
}
