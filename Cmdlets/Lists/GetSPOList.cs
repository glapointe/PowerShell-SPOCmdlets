using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Data.EventReceivers;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Get", "SPOList", SupportsShouldProcess = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Retrieves the collection of lists or a specific list.",
        "If the -List parameter is not provided then all the lists the given Site are returned; otherwise, a single list will be returned matching the identifier provided.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOList -Web \"/\"",
        Remarks = "Retrieves all lists for the root Site.")]
    [Example(Code = "PS C:\\> Get-SPOList -Web \"/\" -Identity \"Shared Documents\"",
        Remarks = "Retrieves a single list for the given Site.")]
    public class GetSPOList : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to retrieve.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the List to retrieve.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind Identity { get; set; }

        [Parameter(
            Position = 2, 
            ValueFromPipeline = false, 
            Mandatory = false,
            HelpMessage = "Retrieve additional details useful for examining a single object.")]
        public SwitchParameter Detail { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            Web web = ctx.Site.OpenWeb(Web.Read());

            if (Identity != null)
            {
                SPOList list = Identity.Read(web, Detail);
                WriteObject(list);
                return;
            }
            var lists = web.Lists;
            ctx.Load(web);
            ctx.Load(lists);
            ctx.ExecuteQuery();

            
            foreach (var list in lists)
            {
                if (Identity == null)
                {
                    SPOList.LoadList(ctx, list, Detail);
                    WriteObject(new SPOList(list));
                    continue;
                }
            }
        }
    }
}
