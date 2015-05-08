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
    [Cmdlet("Add", "SPOContentType", SupportsShouldProcess = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Adds a new content type to an existing List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Add-SPOContentType -Web \"/\" -List \"Documents\" -ContentType \"My Content Type\"",
        Remarks = "This example adds a content type with the name of My Content Type to the Documents list.")]
    public class AddSPOContentType : BaseSPOCmdlet
    {
        public AddSPOContentType()
        {
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to create the new Content Type within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the List to add the Content Type to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 2,
            Mandatory = true,
            HelpMessage = "Sets a value that specifies the Content Type that the new Content Type will inherit from. The type must be a valid name, Content Type ID, Microsoft.SharePoint.Client.ContentType object or SPOContentType object.")]
        public SPOContentTypePipeBind ContentType { get; set; }


        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());



            SPOContentType newContentType = null;
            SPOList list = List.Read(web, false);
            SPOContentType parentContentType = ContentType.Read(web);
            if (parentContentType == null)
            {
                throw new ArgumentException("Unable to locate the specified parent content type.");
            }
            SPOContentType existingContentType = SPOContentType.GetContentType(ctx, list.List.ContentTypes, parentContentType.Name);
            if (existingContentType != null)
            {
                WriteWarning("The content type \"" + parentContentType.Name + "\" already exists within the List.");
                WriteObject(existingContentType);
                return;
            }

            ContentType ct = list.List.ContentTypes.AddExistingContentType(parentContentType.ContentType);
            ctx.ExecuteQuery();
            SPOContentType.LoadContentType(ctx, ct);
            newContentType = new SPOContentType(ct);

            WriteObject(newContentType);
        }

    }
}
