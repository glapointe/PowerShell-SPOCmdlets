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
    [Cmdlet("Remove", "SPOWeb", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [CmdletGroup("Sites")]
    [CmdletDescription("Removes a Site from the current Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOWeb))]
    [Example(Code = "PS C:\\> Remove-SPOWeb -Web \"/childsite\"",
        Remarks = "This example removes the childsite Site from the Site Collection.")]
    public class RemoveSPOWeb : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Alias("Web"),
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to delete.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Identity { get; set; }

        [Parameter(
            Position = 1,
            Mandatory = false,
            HelpMessage = "If the Site contains child Sites then recursively delete all child Sites.")]
        public SwitchParameter Recurse { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            string surl = Identity.Read();
            if (surl == "/")
                throw new InvalidOperationException("You cannot delete the root Site of a Site Collection.");

            if (!base.ShouldProcess(Identity.ToString(), "Remove Web"))
                return;

            Web web = ctx.Site.OpenWeb(surl);
            if (Recurse)
            {
                Delete(web);
            }
            else
            {
                web.DeleteObject();
                web.Context.ExecuteQuery();
            }
        }

        private void Delete(Web web)
        {
            if (!base.ShouldProcess(Identity.ToString()))
                return;

            var ctx = base.Context;
            if (!web.IsObjectPropertyInstantiated("Webs"))
            {
                ctx.Load(web.Webs);
                ctx.ExecuteQuery();
            }
            if (web.Webs.Count > 0)
            {
                foreach (Web childWeb in web.Webs)
                {
                    Delete(childWeb);
                }
                if (web.Webs.Count == 0)
                {
                    web.DeleteObject();
                    web.Context.ExecuteQuery();
                }
            }
            else
            {
                web.DeleteObject();
                web.Context.ExecuteQuery();
            }
        }
    }
}
