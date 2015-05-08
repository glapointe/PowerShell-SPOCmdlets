using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Get", "SPOFolder", SupportsShouldProcess = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Retrieves a specific Folder from a given Site.",
        "Retrieve a specific SPOFolder object from a Site given the server relative URL to the folder. The Site provided by the -Web parameter must correspond to the Site where the folder is located or else an argument out of range exception will be thrown.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOFolder -Web \"/\" -Url \"/Shared Documents/My Folder\"",
        Remarks = "Retrieves the folder \"My Folder\" from the Documents library within the root Site of the current Site Collection.")]
    public class GetSPOFolder : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the File to retrieve.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("Identity", "Folder"),
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "The server relative URL to the folder to retrieve.")]
        public string Url { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            SPOWeb web = new SPOWeb(ctx.Site.OpenWeb(Web.Read()));
            WriteObject(web.GetFolderByServerRelativeUrl(Url));
        }
    }
}
