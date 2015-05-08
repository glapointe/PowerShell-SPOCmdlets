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

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Solutions
{
    [Cmdlet("Add", "SPOSolution", SupportsShouldProcess = false)]
    [CmdletGroup("Solutions")]
    [CmdletDescription("Add a WSP to the Solution Gallery of a Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Add-SPOSolution -LiteralPath \"c:\\customtemplate.wsp\"",
        Remarks = "This example uploads a custom Sandbox Solution to the Solution Gallery of the current Site Collection.")]
    public class AddSPOSolution : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        ValidateFileExists,
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "Specify the path to the solution package (WSP) to add to the Solution Gallery.")]
        public string LiteralPath { get; set; }

        [Parameter(Position = 2, HelpMessage = "Overwrite an existing file if present.")]
        public SwitchParameter Overwrite { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            ctx.Load(ctx.Site, s => s.ServerRelativeUrl);
            ctx.ExecuteQuery();

            FileInfo file = new FileInfo(LiteralPath);
            string leafName = file.Name;
            string url = ctx.Site.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/solutions/" + leafName;
            using (FileStream stream = new FileStream(LiteralPath, FileMode.Open))
            {
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, 
                    url, stream, Overwrite);
            }
            SPOWeb web = new SPOWeb(ctx.Site.RootWeb);
            WriteObject(web.GetFileByServerRelativeUrl(url));
        }
    }
}
