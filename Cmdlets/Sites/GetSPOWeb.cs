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
    [Cmdlet("Get", "SPOWeb", SupportsShouldProcess = false)]
    [CmdletGroup("Sites")]
    [CmdletDescription("Retrieves one or more Sites within the current Site Collection.", "Retrieves a single Site given the server relative URL, GUID, or existing SPOWeb object. If a Site identifier is not provided then all Sites are provided.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOWeb \"/childsite\"",
        Remarks = "This example retrieves the core details for the childsite sub-site of the current Site Collection.")]
    [Example(Code = "PS C:\\> Get-SPOWeb \"/\" -Detail",
        Remarks = "This example retrieves the core details along with additional information for the root site of the current Site Collection.")]
    [Example(Code = "PS C:\\> Get-SPOWeb -Detail",
        Remarks = "This example retrieves the core details along with additional information for all Sites within the current Site Collection.")]
    public class GetSPOWeb : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty, 
        Alias("Web"),
        Parameter(
            Position = 0, 
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the Site to retrieve.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Identity { get; set; }

        [Parameter(
            Position = 1,
            ValueFromPipeline = false,
            Mandatory = false,
            HelpMessage = "Retrieve additional details useful for examining a single object.")]
        public SwitchParameter Detail { get; set; }

        [Parameter(
            Position = 2,
            ValueFromPipeline = false,
            Mandatory = false,
            HelpMessage = "When returning multiple items (-Identity is not provided), include app webs in the resultset.")]
        public SwitchParameter IncludeAppWebs { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Identity == null)
            {
                Web rootWeb = ctx.Site.RootWeb;
                SPOWeb.LoadWeb(ctx, rootWeb, Detail);
                WriteObject(new SPOWeb(rootWeb));
                IterateWebs(ctx, rootWeb, Detail);
            }
            else
            {
                string surl = Identity.Read();
                Web web = ctx.Site.OpenWeb(surl);
                SPOWeb.LoadWeb(ctx, web, Detail);

                WriteObject(new SPOWeb(web));
            }
        }

        private void IterateWebs(CmdletContext ctx, Web parentWeb, bool detailed)
        {
            var childWebs = parentWeb.Webs;
            ctx.Load(childWebs);
            ctx.ExecuteQuery();
            foreach (Web childWeb in childWebs)
            {
                SPOWeb.LoadWeb(ctx, childWeb, detailed);
                SPOWeb web = new SPOWeb(childWeb); 
                if (web.WebTemplate != "APP" || (web.WebTemplate == "APP" && IncludeAppWebs))
                    WriteObject(web);
                IterateWebs(ctx, childWeb, detailed);
            }
        }
    }
}
