using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("New", "SPOListFolder", SupportsShouldProcess = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a Folder in a list or library.",
        "Creates a folder in the specified list. The Site provided by the -Web parameter must correspond to the Site where the folder and List is located or else an argument out of range exception will be thrown.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> New-SPOListFolder -Web \"/\" -List \"Documents\" -ListRelativeUrlFolderPath \"My Top Level Folder/My Sub-Folder\"",
        Remarks = "Creates the folders \"My Top Level Folder/My Sub-Folder\" in the Documents library within the root Site of the current Site Collection.")]
    public class NewSPOListFolder : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to add the item to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List to add the item to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [Parameter(
            Position = 2,
            Mandatory = true,
            HelpMessage = "Create the item within a subfolder inside the list. Only include the path relative to the root of the list.")]
        public string ListRelativeUrlFolderPath { get; set; }
        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            var web = ctx.Site.OpenWeb(Web.Read());
            SPOList list = List.Read(web, true);
            WriteObject(list.GetFolder(ListRelativeUrlFolderPath, true, true));
        }
    }
}
