using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets
{
    [Cmdlet("Connect", "SPOSite", SupportsShouldProcess = false)]
    [CmdletGroup("General")]
    [CmdletDescription("Connects a SharePoint Online user to a SharePoint Online Site Collection. This cmdlet must be run before any other Site Collection scoped SharePoint Online cmdlets can run.",
        "The Connect-SPOSite cmdlet connects a SharePoint Online user to the specified SharePoint Online Site Collection.\r\n\r\nOnly a single SharePoint Online site connection is maintained from any single Windows PowerShell session. In other words, this is a per-site collection user connection. Running the Connect-SPOSite cmdlet twice implicitly disconnects the previous connection. The Windows PowerShell session will be set to serve the new SharePoint Online user and site collection specified. If you are connecting to an on-premises version of SharePoint and not SharePoint Online then use the -NetworkCredentials switch parameter to treat the provided credentials as network credentials.")]
    [RelatedCmdlets(typeof(DisconnectSPOSite))]
    [Example(Code = "PS C:\\> Connect-SPOSite -Url \"https://contoso.sharepoint.com/\" -Credential \"admin@contoso.com\"",
        Remarks = "This example connects to the Site Collection at https://contoso.sharepoint.com/.")]
    [Example(Code = "PS C:\\> $username = \"admin@contoso.sharepoint.com\"\r\n" +
                    "PS C:\\> $password = \"password\"\r\n" +
                    "PS C:\\> $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $(ConvertTo-SecureString $password -AsPlainText -Force)\r\n" +
                    "PS C:\\> Connect-SPOSite -Url \"http://contoso.sharepoint.com/\" -Credential $cred",
        Remarks = "This example connects to the Site Collection at https://contoso.sharepoint.com/ given a specific username and password.")]
    public class ConnectSPOSite : PSCmdlet
    {
        [ValidateNotNullOrEmpty, Parameter(Position = 0, ValueFromPipeline = true, Mandatory = true)]
        public UrlPipeBind Url { get; set; }

        [Parameter(Position = 1, ValueFromPipeline = true)]
        public CredentialPipeBind Credential { get; set; }

        [Parameter(Position = 2, HelpMessage = "If specified then the provided credentials must be network credentials and it is assumed that you are connecting to an on-premises version of SharePoint and not SharePoint Online.")]
        public SwitchParameter NetworkCredentials { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            Uri url = this.Url.Read();
            PSCredential credentials = (this.Credential != null) ? this.Credential.Read(base.Host) : CredentialPipeBind.PromptForCredentials(base.Host, string.Empty);
            if (credentials == null)
            {
                throw new InvalidOperationException("No credentials were provided for the site.");
            }
            DisconnectSPOSite.DisconnectCurrentSite();
            SPOSiteContext.CurrentSiteContext = SPOSiteHelper.InstantiateSPOSiteContext(url, credentials, base.Host, NetworkCredentials);

            WriteObject(SPOSiteContext.CurrentSiteContext);
        }

    }

}
