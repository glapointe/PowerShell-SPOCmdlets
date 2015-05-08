using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets
{
    [Cmdlet("Disconnect", "SPOSite", SupportsShouldProcess = false)]
    [CmdletGroup("General")]
    [CmdletDescription("Disconnects from a SharePoint Online site.", 
        "The Disconnect-SPOSite cmdlet explicitly disconnects a previously created SharePoint Online site connection.\r\n\r\nYou can disconnect the site by running the Disconnect-SPOSite cmdlet or by running the Connect-SPOSite cmdlet to initiate another connection. For more information, see Connect-SPOSite.\r\n\r\nEven after a connection is terminated, operations that were started before the connection is terminated will run to completion. In other words, long-running operations will not be terminated by running the Disconnect-SPOSite cmdlet or the Connect-SPOSite cmdlet.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Disconnect-SPOSite",
        Remarks = "This example explicitly disconnects a previously created SharePoint Online service connection.")]
    public class DisconnectSPOSite : PSCmdlet
    {
        internal static bool DisconnectCurrentSite()
        {
            if (SPOSiteContext.CurrentSiteContext != null)
            {
                if (SPOSiteContext.CurrentSiteContext.Context != null)
                    SPOSiteContext.CurrentSiteContext.Context.Dispose();

                SPOSiteContext.CurrentSiteContext = null;
                return true;
            }
            return false;
        }

        protected override void ProcessRecord()
        {
            if (!DisconnectCurrentSite())
            {
                throw new InvalidOperationException("There is no service currently connected.");
            }
        }

    }
}
