using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Lapointe.SharePointOnline.PowerShell.Data;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Sites
{
    [Cmdlet("Get", "SPOContextSite", SupportsShouldProcess = false)]
    [CmdletGroup("Sites")]
    [CmdletDescription("Retrieves the details for the currently connected Site Collection.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOContextSite",
        Remarks = "This example retrieves the core details for the currently connected site.")]
    [Example(Code = "PS C:\\> Get-SPOContextSite -Detail",
        Remarks = "This example retrieves the core details along with additional information such as upgrade information, event receivers, owner, etc.")]
    public class GetSPOContextSite : BaseSPOCmdlet
    {
        [Parameter(
            Position = 0, 
            ValueFromPipeline = false,
            Mandatory = false,
            HelpMessage = "Retrieve additional details useful for examining a single object.")]
        public SwitchParameter Detail { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            WriteObject(SPOSiteContext.CurrentSiteContext.GetSite(Detail));
        }
    }
}
