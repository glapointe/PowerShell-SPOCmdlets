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
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet("Get", "SPOContentType", SupportsShouldProcess = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Retrieves the Content Types for the Site or List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOContentType -Web \"/\"",
        Remarks = "Retrieves all the Content Types for the root Site of the current Site Collection.")]
    [Example(Code = "PS C:\\> Get-SPOContentType -Web \"/\" -Identity \"Document\"",
        Remarks = "Retrieves the Document Content Types from the root Site of the current Site Collection.")]
    public class GetSPOContentType : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List whose Content Types will be retrieved.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.\r\n\r\nIf not specified then the current Site Collection's Event Receiver Definitions will be retrieved."),
        Parameter(
            ParameterSetName = "Web",
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site whose Content Types will be retrieved.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.\r\n\r\nIf not specified then the current Site Collection's Event Receiver Definitions will be retrieved.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the List whose Content Types will be retrieved.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 2,
            Mandatory = false,
            HelpMessage = "The name or ID of the Content Type to retrieve.")]
        public string Identity { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            string surl = Web.Read();
            Web web = ctx.Site.OpenWeb(surl);

            ContentTypeCollection contentTypes = null;
            if (ParameterSetName == "Web")
            {
                contentTypes = web.ContentTypes;
            }
            else if (ParameterSetName == "List")
            {
                SPOList list = List.Read(web, false);
                contentTypes = list.List.ContentTypes;
            }
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();

            foreach (ContentType ct in contentTypes)
            {
                if (string.IsNullOrEmpty(Identity) || ct.Name.ToLower() == Identity.ToLower() || ct.Id.StringValue.ToLower() == Identity.ToLower())
                {
                    SPOContentType.LoadContentType(ctx, ct);
                    WriteObject(new SPOContentType(ct));
                }
            }
        }

    }
}
