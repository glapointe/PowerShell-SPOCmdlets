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
using System.Xml;
using Lapointe.SharePointOnline.PowerShell.Data.Fields;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Fields
{
    [Cmdlet("Add", "SPOField", SupportsShouldProcess = false)]
    [CmdletGroup("Fields")]
    [CmdletDescription("Adds an existing Field to a List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Add-SPOField -Web \"/\" -List \"Documents\" -Identity \"My Field\" ",
        Remarks = "This example adds the field \"My Field\" to the Documents library.")]
    public class AddSPOField : BaseSPOCmdlet
    {
        public AddSPOField()
        {
        }

        [ValidateNotNullOrEmpty]
        [Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to create the new Content Type within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty]
        [Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the List to add the Content Type to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty]
        [Alias("Name", "Title"),
        Parameter(
          ParameterSetName = "List_ExistingField",
          Position = 2,
          Mandatory = true,
          HelpMessage = "Specifies the existing field to add to the list. The type must be a valid GUID, internal field name, or field title.")]
        public SPOFieldPipeBind Identity { get; set; }


        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());

            SPOField fieldToAdd = Identity.Read(web);
            if (fieldToAdd == null)
            {
                WriteError(new ErrorRecord(new ArgumentOutOfRangeException("The specified field could not be found! The field was not added to the list."), null, ErrorCategory.InvalidData, Identity));
                return;
            }
            SPOList list = List.Read(web, false);
            SPOField existingField = SPOField.GetField(ctx, list.List.Fields, fieldToAdd.InternalName);
            if (existingField != null)
            {
                WriteWarning("The specified field already exists within the list.");
                WriteObject(existingField);
                return;
            }
            Field field = list.List.Fields.Add(fieldToAdd.Field);
            ctx.Load(field);
            ctx.ExecuteQuery();
            WriteObject(new SPOField(field));
        }

    }
}
