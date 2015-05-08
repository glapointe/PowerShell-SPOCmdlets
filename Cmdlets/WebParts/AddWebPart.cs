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
using Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.WebParts
{
    [Cmdlet("Add", "SPOWebPart", SupportsShouldProcess = false, DefaultParameterSetName = "WikiPage")]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Adds a web part to a web part page or wiki page.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOFile))]
    [Example(Code = "PS C:\\> Add-SPOWebPart -Web \"/\" -File \"/default.aspx\" -WebPart \"c:\\announcements.webpart\" -Title \"Announcements\" -Zone \"MiddleColumn\" -ZoneIndex 1",
        Remarks = "This example adds an announcements web part to the default.aspx web part page located in the root web of the root site collection.")]
    [Example(Code = "PS C:\\> Add-SPOWebPart -Web \"/\" -File \"/sitepages/home.aspx\" -WebPart \"c:\\announcements.webpart\" -Row 1 -Column 1 -AddSpace",
        Remarks = "This example adds an announcements web part to the home.aspx wiki page located in the root web of the root site collection.")]
    public class AddWebPart : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the file to add the web part to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the file to add the web part to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name/list_name/file_name, or an SPOFile or SPOWikiPage object, or Microsoft.SharePoint.Client.File object.")]
        public SPOFilePipeBind File { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 2,
            Mandatory = true,
            HelpMessage = "The web part XML file to add to the page.")]
        public XmlDocumentPipeBind WebPart { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 3,
            Mandatory = false,
            HelpMessage = "The title of the web part to add.")]
        public string Title { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "WebPartPage",
            Position = 4, 
            Mandatory = true, 
            HelpMessage = "The zone to add the web part to.")]
        public string Zone { get; set; }

        [Parameter(
            ParameterSetName = "WebPartPage",
            Position = 5,
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        public int ZoneIndex { get; set; }

        [Parameter(
            ParameterSetName = "WikiPage",
            Position = 6,
            Mandatory = true,
            HelpMessage = "The zone to add the web part to.")]
        public int Row { get; set; }

        [Parameter(
            ParameterSetName = "WikiPage",
            Position = 7,
            Mandatory = true,
            HelpMessage = "The zone index to add the web part to.")]
        public int Column { get; set; }

        [Parameter(
            ParameterSetName = "WikiPage",
            Position = 8,
            Mandatory = false,
            HelpMessage = "Add some space before the web part.")]
        public SwitchParameter AddSpace { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            
            string xml = WebPart.Read().OuterXml;
            var web = ctx.Site.OpenWeb(Web.Read());
            SPOFile file = File.Read(web);
            if (ParameterSetName == "WikiPage")
            {
                SPOWikiPage wikiPage = file as SPOWikiPage;
                if (wikiPage == null)
                    throw new Exception("The specified file is either not found or is not a wiki page.");
                WriteObject(wikiPage.AddWebPartToWikiPage(xml, Title, Row, Column, AddSpace));
            }
            else
            {
                WriteObject(file.AddWebPartToWebPartPage(xml, Zone, ZoneIndex, Title));
            }
        }
    }
}
