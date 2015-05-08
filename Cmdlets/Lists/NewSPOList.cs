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
    [Cmdlet("New", "SPOList", SupportsShouldProcess = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new list within an existing Site.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList))]
    [Example(Code = "PS C:\\> New-SPOList -Web \"/\" -Url \"Lists/MyNewList\" -TemplateType 100 -Title \"My New List\"",
        Remarks = "This example creates a new list with the URL name of Lists/MyNewList under the root Site of the current Site Collection.")]
    public class NewSPOList : BaseSPOCmdlet
    {
        public NewSPOList()
        {
            QuickLaunchOption = QuickLaunchOptions.DefaultValue;
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to create the new List within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "The URL name of the List to create. If creating a new List at \"http://contoso.sharepoint.com/site1/Lists/MyList\" then specify a value of \"Lists/MyList\".")]
        public string Url { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("Name"),
        Parameter(
            Position = 2,
            Mandatory = true,
            HelpMessage = "Specifies the title of the new List.")]
        public string Title { get; set; }

        [Parameter(
            Position = 3,
            Mandatory = false,
            HelpMessage = "Describes the new List. If no description is specified, the entry is left blank.")]
        public string Description { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 4,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the List schema of the new List.")]
        public string CustomSchemaXml { get; set; }

        [Parameter(
            Position = 5,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the identifier of the document template for the new List.")]
        public int DocumentTemplateType { get; set; }

        [Parameter(
            Position = 6,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether the new List is displayed on the Quick Launch of the Site.")]
        public QuickLaunchOptions QuickLaunchOption { get; set; }

        [Parameter(
            Position = 7,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the List server template of the new List.")]
        public int TemplateType { get; set; }

        [Parameter(
            Position = 8,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the Feature identifier of the Feature that contains the List schema for the new List.")]
        public Guid TemplateFeatureId { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 9,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the properties of the data source of the new List.")]
        public Dictionary<string, string> DataSourceProperties { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(Web.Read());

            var listci = new ListCreationInformation
            {
                CustomSchemaXml = CustomSchemaXml,
                Description = Description,
                Title = Title,
                Url = Url,
                DocumentTemplateType = DocumentTemplateType,
                DataSourceProperties = DataSourceProperties,
                QuickLaunchOption = QuickLaunchOption,
                TemplateType = TemplateType,
                TemplateFeatureId = TemplateFeatureId
            };

            List newList = web.Lists.Add(listci);
            ctx.ExecuteQuery();
            SPOList.LoadList(ctx, newList, true);
            WriteObject(new SPOList(newList));
        }
    }
}
