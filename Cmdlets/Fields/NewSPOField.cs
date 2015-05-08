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
    [Cmdlet("New", "SPOField", SupportsShouldProcess = false, DefaultParameterSetName = "Web")]
    [CmdletGroup("Fields")]
    [CmdletDescription("Creates a new Field within an existing Site or List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> New-SPOField -Web \"/\" -SchemaXml $schema -Options \"AddFieldInternalNameHint\"",
        Remarks = "This example creates a new site column using the schema definition stored within $schema.")]
    public class NewSPOField : BaseSPOCmdlet
    {
        public NewSPOField()
        {
            Options = AddFieldOptions.DefaultValue;
        }

        [ValidateNotNullOrEmpty]
        [Parameter(
            ParameterSetName = "List",
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to create the new Content Type within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        [Parameter(
            ParameterSetName = "Web",
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to create the new Content Type within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }


        [ValidateNotNullOrEmpty]
        [Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the List to add the Content Type to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Web",
            Position = 2,
            Mandatory = true,
            HelpMessage = "Specifies the XML schema of the new Field.")]
        [Parameter(
            ParameterSetName = "List",
            Position = 2,
            Mandatory = true,
            HelpMessage = "Specifies the XML schema of the new Field.")]
        public XmlDocumentPipeBind SchemaXml { get; set; }

        [Parameter(
            ParameterSetName = "List",
            Position = 3,
            Mandatory = false,
            HelpMessage = "If specified, add the field to the default view.")]
        public SwitchParameter AddToDefaultView { get; set; }


        [Parameter(
            ParameterSetName = "Web",
            Position = 4,
            Mandatory = false,
            HelpMessage = "Field creation options.")]
        [Parameter(
            ParameterSetName = "List",
            Position = 4,
            Mandatory = false,
            HelpMessage = "Field creation options.")]
        public AddFieldOptions Options { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());

            XmlDocument schemaXmlDoc = SchemaXml.Read();
            string xml = schemaXmlDoc.OuterXml;

            if (ParameterSetName == "Web")
            {

                SPOField existingField = SPOField.GetField(ctx, web.AvailableFields, schemaXmlDoc);
                if (existingField != null)
                {
                    WriteWarning("The field " + schemaXmlDoc.DocumentElement.GetAttribute("Name") + " already exists!");
                    WriteObject(existingField);
                    return;
                }
                Field field = web.Fields.AddFieldAsXml(xml, false, Options);
                ctx.Load(field);
                ctx.ExecuteQuery();
                WriteObject(new SPOField(field));
            }
            else if (ParameterSetName == "List")
            {
                SPOList list = List.Read(web, false);
                SPOField existingField = SPOField.GetField(ctx, list.List.Fields, schemaXmlDoc);
                if (existingField != null)
                {
                    WriteWarning("The field " + schemaXmlDoc.DocumentElement.GetAttribute("Name") + " already exists!");
                    WriteObject(existingField);
                    return;
                }
                Field field = list.List.Fields.AddFieldAsXml(xml, AddToDefaultView, Options);
                ctx.Load(field);
                ctx.ExecuteQuery();
                WriteObject(new SPOField(field));
            }
        }

    }
}
