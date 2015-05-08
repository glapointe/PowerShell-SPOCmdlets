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

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Sites
{
    [Cmdlet("New", "SPOWeb", SupportsShouldProcess = false)]
    [CmdletGroup("Sites")]
    [CmdletDescription("Creates a new child Site within an existing Site.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOWeb))]
    [Example(Code = "PS C:\\> New-SPOWeb -ParentWeb \"/\" -Url \"childsite\" -WebTemplate \"STS#0\" -Title \"Child Site\" -UseSamePermissionsAsParentSite",
        Remarks = "This example creates a new child site with the URL name of childsite under the root Site of the current Site Collection.")]
    public class NewSPOWeb : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the parent Site to create the new Site within.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind ParentWeb { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "The URL name of the Site to create. If creating a new Site at \"http://contoso.sharepoint.com/site1/newsite\" then specify a value of \"newsite\".")]
        public string Url { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("Template"),
        Parameter(
            Position = 2,
            Mandatory = false,
            HelpMessage = "Specifies the Web template for the new Site. The template must already exist. If no template is specified, no template is applied and a template can be selected later.")]
        public string WebTemplate { get; set; }

        [ValidateNotNullOrEmpty,
        Alias("Name"),
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "Specifies the title of the new Site.")]
        public string Title { get; set; }

        [Parameter(
            Position = 4,
            Mandatory = false,
            HelpMessage = "Describes the new Site. If no description is specified, the entry is left blank.")]
        public string Description { get; set; }

        [Parameter(
            Position = 5,
            Mandatory = false,
            HelpMessage = "Specifies the language template identifier for the new Site. If no language is specified, the Site is created with the same language that was specified when the product was installed.\r\n\r\nThe type must be a valid language identifier (LCID).")]
        public int Language { get; set; }

        [Parameter(
            Position = 6,
            Mandatory = false,
            HelpMessage = "Specifies that this Site is to be created with unique permissions.")]
        public SwitchParameter UniquePermissions { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            Web web = ctx.Site.OpenWeb(ParentWeb.Read());

            var webci = new WebCreationInformation
            {
                Description = Description,
                Language = Language,
                Title = Title,
                Url = Url,
                UseSamePermissionsAsParentSite = !UniquePermissions,
                WebTemplate = WebTemplate
            };
            Web newWeb = web.Webs.Add(webci);
            ctx.ExecuteQuery();
            SPOWeb.LoadWeb(ctx, newWeb, true);
            WriteObject(new SPOWeb(newWeb));
        }
    }
}
