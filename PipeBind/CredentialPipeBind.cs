using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class CredentialPipeBind : CmdletPipeBind<PSCredential>
    {
        private PSCredential m_credential;
        private string m_userName;

        public CredentialPipeBind(PSCredential inputObject) : base(inputObject)
        {
        }

        public CredentialPipeBind(string inputObject)
        {
            this.m_userName = inputObject;
        }

        protected override void Discover(PSCredential instance)
        {
            if (instance != null)
            {
                this.m_userName = instance.UserName;
                this.m_credential = instance;
            }
            else
            {
                this.m_userName = null;
                this.m_credential = null;
            }
        }

        public static bool IsCredentialValid(PSCredential credential, PSHost host)
        {
            if (host == null)
            {
                throw new ArgumentNullException("host");
            }
            if ((credential == null) || ((credential.Password != null) && (credential.Password.Length != 0)))
            {
                return true;
            }
            host.UI.WriteWarningLine("Password cannot be empty.");
            return false;
        }

        public static PSCredential PromptForCredentials(PSHost host, string userName)
        {
            if (host == null)
            {
                throw new ArgumentNullException("host");
            }
            PSCredential credential = null;
            do
            {
                credential = host.UI.PromptForCredential("SharePoint Online PowerShell", "Enter your credentials.", (credential == null) ? userName : credential.UserName, string.Empty);
            }
            while (!IsCredentialValid(credential, host));
            return credential;
        }

        public override PSCredential Read()
        {
            if (this.m_credential == null)
            {
                throw new InvalidOperationException("Can't read credentials without a PSHost.");
            }
            return this.m_credential;
        }

        public PSCredential Read(PSHost host)
        {
            if ((this.m_credential == null) || !IsCredentialValid(this.m_credential, host))
            {
                this.m_credential = PromptForCredentials(host, this.m_userName);
            }
            return this.m_credential;
        }
    }

}
