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
using Lapointe.SharePointOnline.PowerShell.Data.Fields;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.ContentTypes
{
    [Cmdlet("New", "SPOContentType", SupportsShouldProcess = false)]
    [CmdletGroup("Content Types")]
    [CmdletDescription("Creates a new content type within an existing Site or List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> New-SPOContentType -Web \"/\" -Name \"My Content Type\" -ParentContentType $contentType",
        Remarks = "This example creates a new content type with the name of My Content Type under the root Site of the current Site Collection.")]
    public class NewSPOContentType : BaseSPOCmdlet
    {
        public NewSPOContentType()
        {
        }

        [ValidateNotNullOrEmpty]
        [Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to create the new Content Type within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("Title"),
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "Specifies the name of the new Content Type.")]
        public string Name { get; set; }

        [Parameter(
            Position = 2,
            Mandatory = false,
            HelpMessage = "Describes the new Content Type. If no description is specified, the entry is left blank.")]
        public string Description { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 3,
            Mandatory = false,
            HelpMessage = "Describes the new Content Type. If no description is specified, the entry is left blank.")]
        public string Group { get; set; }


        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 4,
            Mandatory = true,
            HelpMessage = "Sets a value that specifies the Content Type that the new Content Type will inherit from. The type must be a valid name, Content Type ID, Microsoft.SharePoint.Client.ContentType object or SPOContentType object.")]
        public SPOContentTypePipeBind ParentContentType { get; set; }


        [Parameter(
            Position = 5,
            Mandatory = false,
            HelpMessage = "An array of internal field names to add to the content type.")]
        public string[] FieldsToAdd { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());



            SPOContentType newContentType = null;
            SPOContentType parentContentType = ParentContentType.Read(web);
            if (parentContentType == null)
            {
                throw new ArgumentException("Unable to locate the specified parent content type.");
            }
            SPOContentType existingContentType = SPOContentType.GetContentType(ctx, web.AvailableContentTypes, Name);
            if (existingContentType != null)
            {
                WriteWarning("The content type \"" + Name + "\" already exists within the Site.");
                WriteObject(existingContentType);
                return;
            }
            List<SPOField> fields = new List<SPOField>();
            if (FieldsToAdd != null)
            {
                foreach (string fieldName in FieldsToAdd)
                {
                    SPOField existingField = SPOField.GetField(ctx, web.AvailableFields, fieldName);
                    if (existingField == null)
                        WriteError(new ErrorRecord(new ArgumentOutOfRangeException("Unable to locate field " + fieldName + ". Content Type will not be created."), null, ErrorCategory.InvalidData, web.AvailableFields));
                    else
                        fields.Add(existingField);
                }
                if (fields.Count != FieldsToAdd.Length)
                    return;
            }

            var ctli = new ContentTypeCreationInformation();
            ctli.Description = Description;
            ctli.Group = Group;
            ctli.Name = Name;
            ctli.ParentContentType = parentContentType.ContentType;

            ContentType ct = web.ContentTypes.Add(ctli);
            ctx.ExecuteQuery();
            SPOContentType.LoadContentType(ctx, ct);
            newContentType = new SPOContentType(ct);

            foreach (SPOField field in fields)
            {
                FieldLinkCreationInformation flci = new FieldLinkCreationInformation();
                flci.Field = field.Field;
                newContentType.ContentType.FieldLinks.Add(flci);
            }
            if (fields.Count > 0)
            {
                newContentType.ContentType.Update(true);
                ctx.ExecuteQuery();
            }


            WriteObject(newContentType);
        }


        
    }
}
