using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Management.Automation.Host;
using System.Reflection;
using System.Text;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell
{
    public sealed class CmdletContext : ClientContext
    {
        private PSHost m_powerShellHost;
        private const string USER_AGENT_STRING_FORMAT = "SharePoint Online PowerShell by Gary Lapointe ({0})";

        internal CmdletContext(string webFullUrl, PSHost host) : base(webFullUrl)
        {
            this.Host = host;
        }

        internal CmdletContext(Uri webFullUrl, PSHost host) : base(webFullUrl)
        {
            this.Host = host;
        }

        internal static string GetUserAgent()
        {
            return string.Format(CultureInfo.InvariantCulture, USER_AGENT_STRING_FORMAT, new object[] { ModuleVersion.FullBuildBase });
        }

        internal PSHost Host
        {
            get
            {
                return this.m_powerShellHost;
            }
            private set
            {
                if (value == null)
                {
                    throw new ArgumentNullException("Host");
                }
                this.m_powerShellHost = value;
            }
        }

        private class ModuleVersion
        {
            static ModuleVersion()
            {
                FullBuildVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
                FullBuildBase = new Version(FullBuildVersion).ToString(3) + ".0";
            }
            public static readonly string FullBuildVersion;
            public static readonly string FullBuildBase;
        }

 

    }
}
