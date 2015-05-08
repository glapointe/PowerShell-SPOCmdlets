using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Enums;
using Lapointe.SharePointOnline.PowerShell.Exceptions;
using Lapointe.SharePointOnline.PowerShell.PipeBind;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Set", "SPOWikiPageLayout", SupportsShouldProcess = false, DefaultParameterSetName = "List")]
    [CmdletGroup("Wiki Pages")]
    [CmdletDescription("Sets the layout for a wiki page within an existing List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList))]
    [Example(Code = "PS C:\\> Set-SPOWikiPageLayout -Web \"/\" -List \"Site Pages\" -WikiPageName \"MyWikiPage.aspx\" -WikiPageLayout \"OneColumnSideBar\"",
        Remarks = "This example sets the layout of an existing wiki page within the Site Pages list under the root Site of the root Site Collection.")]
    public class SetSPOWikiPageLayout : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the file to update.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List containing the file to update.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Folder",
            Position = 2,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the Folder containing the file to update.\r\n\r\nThe type must be a valid SPOFolder object.")]
        public SPOFolder Folder { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "The name of the wiki page to update.")]
        public string WikiPageName { get; set; }

        [Parameter(
            Position = 4, 
            Mandatory = true, 
            HelpMessage = "The page layout to set the wiki page to.")]
        public WikiPageLayout WikiPageLayout { get; set; }
        
        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            
            SPOFile page = null;
            if (ParameterSetName == "List")
            {
                var web = ctx.Site.OpenWeb(Web.Read());
                SPOList list = List.Read(web, true);
                page = list.GetRootFolder(true).GetFile(WikiPageName);
            }
            else
            {
                page = Folder.GetFile(WikiPageName);
            }
            if (page != null)
            {
                SPOWikiPage wikiPage = new SPOWikiPage(page.File);
                wikiPage.SetWikiPageLayout(WikiPageLayout);
            }
        }
    }
}
