using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Data.EventReceivers;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.EventReceivers
{
    [Cmdlet("Get", "SPOEventReceiver", SupportsShouldProcess = false)]
    [CmdletGroup("Event Receivers")]
    [CmdletDescription("Retrieves the Event Receivers for the Site Collection or a specific Site.",
        "If the -Web parameter is not provided then the Site Collection scoped Event Receivers are returned; otherwise, the Site scoped Event Receivers for a given Site are returned.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Get-SPOEventReceiver",
        Remarks = "Retrieves all Site Collection scoped Event Receivers.")]
    [Example(Code = "PS C:\\> Get-SPOEventReceiver -Web \"/\"",
        Remarks = "Retrieves all Site scoped Event Receivers for the root Site.")]
    public class GetSPOEventReceiver : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the Site whose Event Receiver Definitions will be retrieved.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.\r\n\r\nIf not specified then the current Site Collection's Event Receiver Definitions will be retrieved.")]
        public SPOWebPipeBind Web { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            if (Web == null)
            {
                var eventReceivers = ctx.Site.EventReceivers;
                ctx.Load(eventReceivers);
                ctx.ExecuteQuery();
                foreach (EventReceiverDefinition eventReceiver in eventReceivers)
                    WriteObject(new SPOEventReceiverDefinition(eventReceiver, ctx.Site.EventReceivers));
            }
            else
            {
                string surl = Web.Read();
                Web web = ctx.Site.OpenWeb(surl);
                var eventReceivers = web.EventReceivers;
                ctx.Load(eventReceivers);
                ctx.ExecuteQuery();

                foreach (EventReceiverDefinition eventReceiver in eventReceivers)
                    WriteObject(new SPOEventReceiverDefinition(eventReceiver, web.EventReceivers));
            }
        }

    }
}
