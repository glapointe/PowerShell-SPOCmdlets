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
using System.Xml;
using System.Security;
using System.Xml.Linq;
using System.Web.UI.WebControls.WebParts;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.WebParts
{
    [Cmdlet("Add", "SPOListViewWebPart", SupportsShouldProcess = false, DefaultParameterSetName = "WikiPage")]
    [CmdletGroup("Web Parts")]
    [CmdletDescription("Adds a web part to a web part page or wiki page.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOFile))]
    [Example(Code = "PS C:\\> Add-SPOListViewWebPart -Web \"/\" -File \"/default.aspx\" -WebPart \"c:\\announcements.webpart\" -Title \"Announcements\" -Zone \"MiddleColumn\" -ZoneIndex 1",
        Remarks = "This example adds an announcements web part to the default.aspx web part page located in the root web of the root site collection.")]
    [Example(Code = "PS C:\\> Add-SPOListViewWebPart -Web \"/\" -File \"/sitepages/home.aspx\" -WebPart \"c:\\announcements.webpart\" -Row 1 -Column 1 -AddSpace",
        Remarks = "This example adds an announcements web part to the home.aspx wiki page located in the root web of the root site collection.")]
    public class AddListViewWebPart : BaseSPOCmdlet
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

        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            Position = 1,
            HelpMessage = "The list to whose view will be added to the page.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        [ValidateNotNull]
        public SPOListPipeBind List { get; set; }

        [Parameter(Mandatory = false,
            HelpMessage = "The name of the view to use for the list view web part.")]
        [ValidateNotNullOrEmpty]
        public string ViewTitle { get; set; }


        [Parameter(Mandatory = false,
            HelpMessage = "The view schema to apply.")]
        [ValidateNotNullOrEmpty]
        public XmlDocumentPipeBind ViewSchema { get; set; }

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


        [Parameter(
            Position = 9,
            Mandatory = false,
            HelpMessage = "Set a specific JSLink file.")]
        public string JSLink { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "The chrome settings for the web part.")]
        public PartChromeType ChromeType
        {
            get
            {
                if (Fields["ChromeType"] == null)
                    return PartChromeType.Default;
                return (PartChromeType)Fields["ChromeType"];
            }
            set { Fields["ChromeType"] = value; }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            XmlDocument viewSchema = null;
            if (ViewSchema != null)
                viewSchema = ViewSchema.Read();
            if (string.IsNullOrEmpty(ViewTitle) && viewSchema == null)
                throw new ArgumentException("You must specify a value for at least one of the ViewTitle or ViewSchema parameters.");
            var web = ctx.Site.OpenWeb(Web.Read());
            SPOFile file = File.Read(web);
            SPOList list = List.Read(web, true);
            Guid viewId = Guid.Empty;
            if (!string.IsNullOrEmpty(ViewTitle))
            {
                var view = list.GetViews().Where(v => v.Title.ToLowerInvariant() == ViewTitle.ToLowerInvariant()).FirstOrDefault();
                if (view == null)
                    throw new ArgumentException("The value specified for the ViewTitle parameter does not match an existing view.");
            }

            string xml = GetXsltListViewWebPartXml(list.Id, list.Url, viewId, viewSchema, file.ServerRelativeUrl, Title, JSLink, ChromeType);
            if (list.BaseTemplate == 106)
            {
                if (viewSchema == null)
                {
                    var view = list.GetViews().Where(v => v.Title.ToLowerInvariant() == ViewTitle.ToLowerInvariant()).FirstOrDefault();
                    viewSchema = new XmlDocument();
                    viewSchema.LoadXml(view.ListViewXml);
                }
                xml = GetListViewWebPartXml(list.Id, list.Url, viewSchema, file.ServerRelativeUrl, Title, ChromeType);
            }
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
        private string GetListViewWebPartXml(Guid listId, string listUrl, XmlDocument viewSchema, string pageUrl, string title, PartChromeType chromeType)
        {
            StringBuilder wp = new StringBuilder(100);

            wp.Append("<WebPart xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://schemas.microsoft.com/WebPart/v2\">");
            wp.AppendFormat("    <Title>{0}</Title>", title);
            wp.AppendFormat("    <FrameType>{0}</FrameType>", chromeType);
            wp.Append("    <Description></Description>");
            wp.Append("    <IsIncluded>true</IsIncluded>");
            wp.Append("    <PartOrder>1</PartOrder>");
            wp.Append("    <FrameState>Normal</FrameState>");
            wp.Append("    <Height />");
            wp.Append("    <Width />");
            wp.Append("    <AllowRemove>true</AllowRemove>");
            wp.Append("    <AllowZoneChange>true</AllowZoneChange>");
            wp.Append("    <AllowMinimize>true</AllowMinimize>");
            wp.Append("    <AllowConnect>true</AllowConnect>");
            wp.Append("    <AllowEdit>true</AllowEdit>");
            wp.Append("    <AllowHide>true</AllowHide>");
            wp.Append("    <IsVisible>true</IsVisible>");
            wp.AppendFormat("    <DetailLink>{0}</DetailLink>", listUrl);
            wp.Append("    <HelpLink />");
            wp.Append("    <HelpMode>Modeless</HelpMode>");
            wp.Append("    <Dir>Default</Dir>");
            wp.Append("    <PartImageSmall />");
            wp.Append("    <MissingAssembly>Cannot import this Web Part.</MissingAssembly>");
            wp.Append("    <PartImageLarge>/_layouts/15/images/itevent.png?rev=38</PartImageLarge>");
            wp.Append("    <IsIncludedFilter />");
            wp.Append("    <Assembly>Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>");
            wp.Append("    <TypeName>Microsoft.SharePoint.WebPartPages.ListViewWebPart</TypeName>");
            wp.Append("    <ExportControlledProperties>false</ExportControlledProperties>");
            wp.Append("    <ConnectionID>00000000-0000-0000-0000-000000000000</ConnectionID>");
            wp.Append("    <WebId xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">00000000-0000-0000-0000-000000000000</WebId>");
            wp.Append("    <ListViewXml xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">");
            if (viewSchema != null)
            {
                viewSchema.DocumentElement.SetAttribute("Name", Guid.NewGuid().ToString("B"));
                viewSchema.DocumentElement.SetAttribute("Url", pageUrl);
                viewSchema.DocumentElement.SetAttribute("Hidden", "TRUE");
                viewSchema.DocumentElement.SetAttribute("DisplayName", "");
                string encodedViewSchema = new XText(viewSchema.OuterXml).ToString();
                wp.Append(encodedViewSchema);
                //wp.Append("&lt;View Name=\"{535F259F-0C8E-4A7A-98A0-A4FCA5D222E4}\" MobileView=\"TRUE\" Type=\"CALENDAR\" Hidden=\"TRUE\" TabularView=\"FALSE\" RecurrenceRowset=\"TRUE\" DisplayName=\"\" Url=\"/sites/sandbox/SitePages/Home.aspx\" Level=\"1\" BaseViewID=\"2\" ContentTypeID=\"0x\" MobileUrl=\"_layouts/15/mobile/viewdaily.aspx\" ImageUrl=\"/_layouts/15/images/events.png?rev=38\"&gt;&lt;Toolbar Type=\"None\" /&gt;&lt;ViewHeader /&gt;&lt;ViewBody /&gt;&lt;ViewFooter /&gt;&lt;ViewEmpty /&gt;&lt;ParameterBindings&gt;&lt;ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /&gt;&lt;ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /&gt;&lt;/ParameterBindings&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"EventDate\" /&gt;&lt;FieldRef Name=\"EndDate\" /&gt;&lt;FieldRef Name=\"fRecurrence\" /&gt;&lt;FieldRef Name=\"EventType\" /&gt;&lt;FieldRef Name=\"WorkspaceLink\" /&gt;&lt;FieldRef Name=\"Title\" /&gt;&lt;FieldRef Name=\"Location\" /&gt;&lt;FieldRef Name=\"Description\" /&gt;&lt;FieldRef Name=\"Workspace\" /&gt;&lt;FieldRef Name=\"MasterSeriesItemID\" /&gt;&lt;FieldRef Name=\"fAllDayEvent\" /&gt;&lt;/ViewFields&gt;&lt;ViewData&gt;&lt;FieldRef Name=\"Title\" Type=\"CalendarMonthTitle\" /&gt;&lt;FieldRef Name=\"Title\" Type=\"CalendarWeekTitle\" /&gt;&lt;FieldRef Name=\"Location\" Type=\"CalendarWeekLocation\" /&gt;&lt;FieldRef Name=\"Title\" Type=\"CalendarDayTitle\" /&gt;&lt;FieldRef Name=\"Location\" Type=\"CalendarDayLocation\" /&gt;&lt;/ViewData&gt;&lt;Query&gt;&lt;Where&gt;&lt;DateRangesOverlap&gt;&lt;FieldRef Name=\"EventDate\" /&gt;&lt;FieldRef Name=\"EndDate\" /&gt;&lt;FieldRef Name=\"RecurrenceID\" /&gt;&lt;Value Type=\"DateTime\"&gt;&lt;Month /&gt;&lt;/Value&gt;&lt;/DateRangesOverlap&gt;&lt;/Where&gt;&lt;/Query&gt;&lt;/View&gt;");
            }
            wp.Append("    </ListViewXml>");
            wp.AppendFormat("    <ListName xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">{0}</ListName>", listId.ToString("B").ToUpperInvariant());
            wp.AppendFormat("    <ListId xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">{0}</ListId>", listId.ToString());
            //wp.Append("    <ViewFlag xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">8921097</ViewFlag>");
            //wp.Append("    <ViewFlags xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">Html Hidden RecurrenceRowset Calendar Mobile</ViewFlags>");
            wp.Append("    <ViewContentTypeId xmlns=\"http://schemas.microsoft.com/WebPart/v2/ListView\">0x</ViewContentTypeId>");
            wp.Append("</WebPart>");
            return wp.ToString();
        }
        private string GetXsltListViewWebPartXml(Guid listId, string listUrl, Guid viewId, XmlDocument viewSchema, string pageUrl, string title, string jslink, PartChromeType chromeType)
        {
            StringBuilder wp = new StringBuilder(100);
            wp.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            wp.Append("<webParts>");
            wp.Append("	<webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">");
            wp.Append("		<metaData>");
            wp.Append("			<type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />");
            wp.Append("			<importErrorMessage>Cannot import this Web Part.</importErrorMessage>");
            wp.Append("		</metaData>");
            wp.Append("		<data>");
            wp.Append("			<properties>");
            wp.Append("				<property name=\"ShowWithSampleData\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"Default\" type=\"string\" />");
            wp.Append("				<property name=\"NoDefaultStyle\" type=\"string\" null=\"true\" />");
            wp.Append("				<property name=\"CacheXslStorage\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"ViewContentTypeId\" type=\"string\" />");
            wp.Append("				<property name=\"XmlDefinitionLink\" type=\"string\" />");
            wp.Append("				<property name=\"ManualRefresh\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"ListUrl\" type=\"string\" />");
            wp.AppendFormat("				<property name=\"ListId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">{0}</property>", listId.ToString());
            if (viewId != Guid.Empty)
                wp.AppendFormat("				<property name=\"ViewGuid\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">{0}</property>", viewId.ToString("B").ToUpperInvariant());
            wp.AppendFormat("				<property name=\"TitleUrl\" type=\"string\">{0}</property>", listUrl);
            wp.Append("				<property name=\"EnableOriginalValue\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"Direction\" type=\"direction\">NotSet</property>");
            wp.Append("				<property name=\"ServerRender\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"ViewFlags\" type=\"Microsoft.SharePoint.SPViewFlags, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">None</property>");
            wp.Append("				<property name=\"AllowConnect\" type=\"bool\">True</property>");
            wp.AppendFormat("				<property name=\"ListName\" type=\"string\">{0}</property>", listId.ToString("B").ToUpperInvariant());
            wp.Append("				<property name=\"ListDisplayName\" type=\"string\" />");
            wp.Append("				<property name=\"AllowZoneChange\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"ChromeState\" type=\"chromestate\">Normal</property>");
            wp.Append("				<property name=\"DisableSaveAsNewViewButton\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"ViewFlag\" type=\"string\" />");
            wp.Append("				<property name=\"DataSourceID\" type=\"string\" />");
            wp.Append("				<property name=\"ExportMode\" type=\"exportmode\">All</property>");
            wp.Append("				<property name=\"AutoRefresh\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"FireInitialRow\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"AllowEdit\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"Description\" type=\"string\" />");
            wp.Append("				<property name=\"HelpMode\" type=\"helpmode\">Modeless</property>");
            wp.Append("				<property name=\"BaseXsltHashKey\" type=\"string\" null=\"true\" />");
            wp.Append("				<property name=\"AllowMinimize\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"CacheXslTimeOut\" type=\"int\">86400</property>");
            wp.AppendFormat("				<property name=\"ChromeType\" type=\"chrometype\">{0}</property>", chromeType.ToString());
            wp.Append("				<property name=\"Xsl\" type=\"string\" null=\"true\" />");
            if (string.IsNullOrEmpty(jslink))
                wp.Append("				<property name=\"JSLink\" type=\"string\" null=\"true\" />");
            else
                wp.AppendFormat("				<property name=\"JSLink\" type=\"string\">{0}</property>", jslink);
            wp.Append("				<property name=\"CatalogIconImageUrl\" type=\"string\">/_layouts/15/images/itgen.png?rev=26</property>");
            wp.Append("				<property name=\"SampleData\" type=\"string\" null=\"true\" />");
            wp.Append("				<property name=\"UseSQLDataSourcePaging\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"TitleIconImageUrl\" type=\"string\" />");
            wp.Append("				<property name=\"PageSize\" type=\"int\">-1</property>");
            wp.Append("				<property name=\"ShowTimelineIfAvailable\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"Width\" type=\"string\" />");
            wp.Append("				<property name=\"DataFields\" type=\"string\" />");
            wp.Append("				<property name=\"Hidden\" type=\"bool\">False</property>");
            wp.AppendFormat("				<property name=\"Title\" type=\"string\">{0}</property>", title);
            wp.Append("				<property name=\"PageType\" type=\"Microsoft.SharePoint.PAGETYPE, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">PAGE_NORMALVIEW</property>");
            wp.Append("				<property name=\"DataSourcesString\" type=\"string\" />");
            wp.Append("				<property name=\"AllowClose\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"InplaceSearchEnabled\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"WebId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">00000000-0000-0000-0000-000000000000</property>");
            wp.Append("				<property name=\"Height\" type=\"string\" />");
            wp.Append("				<property name=\"GhostedXslLink\" type=\"string\">main.xsl</property>");
            wp.Append("				<property name=\"DisableViewSelectorMenu\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"DisplayName\" type=\"string\" />");
            wp.Append("				<property name=\"IsClientRender\" type=\"bool\">False</property>");
            if (viewSchema != null)
            {
                viewSchema.DocumentElement.SetAttribute("Name", Guid.NewGuid().ToString("B"));
                viewSchema.DocumentElement.SetAttribute("Url", pageUrl);
                viewSchema.DocumentElement.SetAttribute("Hidden", "TRUE");
                viewSchema.DocumentElement.SetAttribute("DisplayName", "");
                string encodedViewSchema = new XText(viewSchema.OuterXml).ToString();
                wp.Append("				<property name=\"XmlDefinition\" type=\"string\">");
                wp.Append(encodedViewSchema);
                //wp.Append(string.Format("&lt;View Name=\"{1}\" Type=\"HTML\" Hidden=\"TRUE\" ReadOnly=\"TRUE\" OrderedView=\"TRUE\" DisplayName=\"\" Url=\"{0}\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" &gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=\"TileOrder\" Ascending=\"TRUE\"/&gt;&lt;FieldRef Name=\"Modified\" Ascending=\"FALSE\"/&gt;&lt;/OrderBy&gt;&lt;/Query&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"Title\"/&gt;&lt;FieldRef Name=\"BackgroundImageLocation\"/&gt;&lt;FieldRef Name=\"Description\"/&gt;&lt;FieldRef Name=\"LinkLocation\"/&gt;&lt;FieldRef Name=\"LaunchBehavior\"/&gt;&lt;FieldRef Name=\"BackgroundImageClusterX\"/&gt;&lt;FieldRef Name=\"BackgroundImageClusterY\"/&gt;&lt;/ViewFields&gt;&lt;RowLimit Paged=\"TRUE\"&gt;30&lt;/RowLimit&gt;&lt;JSLink&gt;sp.ui.tileview.js&lt;/JSLink&gt;&lt;XslLink Default=\"TRUE\"&gt;main.xsl&lt;/XslLink&gt;&lt;Toolbar Type=\"Standard\"/&gt;&lt;/View&gt;", pageUrl, ("{" + Guid.NewGuid().ToString() + "}")));
                wp.Append("				</property>");
            }
            wp.Append("				<property name=\"InitialAsyncDataFetch\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"AllowHide\" type=\"bool\">True</property>");
            wp.Append("				<property name=\"ParameterBindings\" type=\"string\">");
            wp.Append("  &lt;ParameterBinding Name=\"dvt_sortdir\" Location=\"Postback;Connection\"/&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"dvt_sortfield\" Location=\"Postback;Connection\"/&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"dvt_startposition\" Location=\"Postback\" DefaultValue=\"\"/&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"dvt_firstrow\" Location=\"Postback;Connection\"/&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"OpenMenuKeyAccessible\" Location=\"Resource(wss,OpenMenuKeyAccessible)\" /&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"open_menu\" Location=\"Resource(wss,open_menu)\" /&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"select_deselect_all\" Location=\"Resource(wss,select_deselect_all)\" /&gt;");
            wp.Append("            &lt;ParameterBinding Name=\"idPresEnabled\" Location=\"Resource(wss,idPresEnabled)\" /&gt;&lt;ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /&gt;&lt;ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /&gt;</property>");
            wp.Append("				<property name=\"DataSourceMode\" type=\"Microsoft.SharePoint.WebControls.SPDataSourceMode, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">List</property>");
            wp.Append("				<property name=\"AutoRefreshInterval\" type=\"int\">60</property>");
            wp.Append("				<property name=\"AsyncRefresh\" type=\"bool\">False</property>");
            wp.Append("				<property name=\"HelpUrl\" type=\"string\" />");
            wp.Append("				<property name=\"MissingAssembly\" type=\"string\">Cannot import this Web Part.</property>");
            wp.Append("				<property name=\"XslLink\" type=\"string\" null=\"true\" />");
            wp.Append("				<property name=\"SelectParameters\" type=\"string\" />");
            wp.Append("			</properties>");
            wp.Append("		</data>");
            wp.Append("	</webPart>");
            wp.Append("</webParts>");
            return wp.ToString();
        }

    }
}
