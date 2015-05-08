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
using Lapointe.SharePointOnline.PowerShell.Data.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Remove", "SPOList", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Removes a List from the specified Site.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList))]
    [Example(Code = "PS C:\\> Remove-SPOWeb -Web \"/childsite\"",
        Remarks = "This example removes the childsite Site from the Site Collection.")]
    public class RemoveSPOList : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to delete.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("List"),
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the List to delete.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind Identity { get; set; }

        [Parameter(
            Position = 2,
            HelpMessage = "If specified the List will be recycled instead of deleted.")]
        public SwitchParameter Recycle { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            if (!base.ShouldProcess(Identity.ToString(), "Remove List"))
                return;

            Web web = ctx.Site.OpenWeb(Web.Read());
            SPOList list = Identity.Read(web, false);
            if (list != null)
            {
                if (Recycle)
                    list.List.Recycle();
                else
                    list.List.DeleteObject();

                list.List.Context.ExecuteQuery();
            }
            else
                throw new ArgumentException("Unable to locate the specified list.");
        }

    }
}
