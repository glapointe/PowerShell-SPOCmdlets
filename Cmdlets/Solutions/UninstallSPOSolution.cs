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
    // Only works on public facing o365 sites so comment out for now.
    /*
    [Cmdlet("Uninstall", "SPOSolution", SupportsShouldProcess = false)]
    [CmdletGroup("Solutions")]
    [CmdletDescription("Add a WSP to the Solution Gallery of a Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Uninstall-SPOSolution -Identity \"customtemplate.wsp\"",
        Remarks = "This example deactivates a custom Sandbox Solution in the Solution Gallery of the current Site Collection.")]
    public class UninstallSPOSolution : BaseSPOCmdlet
    {
        public UninstallSPOSolution() { MajorVersion = 1; }

        [ValidateNotNullOrEmpty,
        Alias("PackageName", "Name"),
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "Specify the name of the solution package (WSP) to deactivate.")]
        public string Identity { get; set; }

        [Parameter(Position = 2, HelpMessage = "Sets the major version of the WSP to deactivate. The default value is 1.")]
        public int MajorVersion { get; set; }

        [Parameter(Position = 3, HelpMessage = "Sets the minor version of the WSP to deactivate. The default value is 0.")]
        public int MinorVersion { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            DesignPackageInfo dpi = new DesignPackageInfo();
            dpi.PackageGuid = Guid.Empty;
            dpi.MajorVersion = MajorVersion;
            dpi.MinorVersion = MinorVersion;
            dpi.PackageName = Identity;

            DesignPackage.UnInstall(ctx, ctx.Site, dpi);
            ctx.ExecuteQuery();


        }
    }
     * */
}
