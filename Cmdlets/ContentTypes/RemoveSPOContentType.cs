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
using Lapointe.SharePointOnline.PowerShell.Data.ContentTypes;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet("Remove", "SPOContentType", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Adds a new content type to an existing List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Remove-SPOContentType -Web \"/\" -List \"Documents\" -ContentType \"My Content Type\"",
        Remarks = "This example removes a content type with the name of My Content Type from the Documents list.")]
    public class RemoveSPOContentType : BaseSPOCmdlet
    {
        public RemoveSPOContentType()
        {
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the Content Type or List containing the Content Type to remove the Content Type from.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List to remove the Content Type from.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 2,
            Mandatory = true,
            HelpMessage = "Sets a value that specifies the Content Type to remove. The type must be a valid name, Content Type ID, Microsoft.SharePoint.Client.ContentType object or SPOContentType object.")]
        public SPOContentTypePipeBind ContentType { get; set; }


        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());

            SPOContentType contentType = null;
            if (List == null)
            {
                contentType = ContentType.Read(web);
            }
            else
            {
                SPOList list = List.Read(web, false);
                contentType = ContentType.Read(list.List); 
            }

            if (contentType == null)
            {
                throw new ArgumentException("Unable to locate the specified content type.");
            }

            if (base.ShouldProcess(contentType.Name, "Remove Content Type"))
            {
                contentType.Delete();
            }
        }

    }
}
