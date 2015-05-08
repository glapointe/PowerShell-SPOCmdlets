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
    [Cmdlet("Set", "SPOWebTheme", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium)]
    [CmdletGroup("Sites")]
    [CmdletDescription("Applies a theme with the specified components to the specified Site.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOWeb))]
    [Example(Code = "PS C:\\> Set-SPOWebTheme -Web \"/\" -ColorPaletteUrl \"/_catalogs/theme/15/palette001.spcolor\"",
        Remarks = "This example sets the theme of the root Site of the current Site Collection.")]
    public class SetSPOWebTheme : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Alias("Web"),
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to apply the theme to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Identity { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            Mandatory = true,
            HelpMessage = "The server relative URL to the color palette file. For example: \"/_catalogs/theme/15/palette001.spcolor\".")]
        public string ColorPaletteUrl { get; set; }

        [Parameter(
            Position = 2,
            Mandatory = false,
            HelpMessage = "The server relative URL to the font file. For example: \"/_catalogs/theme/15/fontscheme001.spfont\".")]
        public string FontSchemeUrl { get; set; }

        [Parameter(
            Position = 3,
            Mandatory = false,
            HelpMessage = "The server relative URL to the background image file. For example: \"/Images/bg.png\".")]
        public string BackgroundImageUrl { get; set; }

        [Parameter(
            Position = 4,
            Mandatory = false,
            HelpMessage = "If specified, store the generated theme files in the root Site, otherwise store them in this Site.")]
        public bool ShareGenerated { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            string surl = Identity.Read();

            if (!base.ShouldProcess(Identity.ToString()))
                return;

            SPOWeb web = new SPOWeb(ctx.Site.OpenWeb(surl));
            web.ApplyTheme(ColorPaletteUrl, FontSchemeUrl, BackgroundImageUrl, ShareGenerated);
        }
    }
}
