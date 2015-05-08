using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.Win32;
using System.Net;

namespace Lapointe.SharePointOnline.PowerShell
{
    internal class SPOSiteHelper
    {
        internal static string HEADER_SHAREPOINT_VERSION = "MicrosoftSharePointTeamServices";

        internal static SPOSiteContext InstantiateSPOSiteContext(Uri url, PSCredential credentials, PSHost host)
        {
            return InstantiateSPOSiteContext(url, credentials, host, false);
        }
        internal static SPOSiteContext InstantiateSPOSiteContext(Uri url, PSCredential credentials, PSHost host, bool isNetworkCredentials)
        {
            if (!IsValidServerVersion(url))
            {
                throw new InvalidOperationException("Could not connect to SharePoint Online: unsupported version of service.");
            }
            CmdletContext context = new CmdletContext(url.AbsoluteUri, host);
            if (isNetworkCredentials)
            {
                context.AuthenticationMode = ClientAuthenticationMode.Default;
                context.Credentials = credentials.GetNetworkCredential();
            }
            else
            {
                context.Credentials = new SharePointOnlineCredentials(credentials.UserName, credentials.Password);
            }

            if (context.HasPendingRequest)
                context.ExecuteQuery();
            return new SPOSiteContext(context);
        }

        internal static bool IsValidServerVersion(Uri url)
        {
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.UserAgent] = CmdletContext.GetUserAgent();
                string str = null;
                try
                {
                    client.DownloadData(url);
                    str = client.ResponseHeaders[HEADER_SHAREPOINT_VERSION];
                }
                catch (WebException exception)
                {
                    if ((exception != null) && (exception.Response != null))
                    {
                        str = exception.Response.Headers[HEADER_SHAREPOINT_VERSION];
                    }
                }
                if (str == null)
                {
                    throw new InvalidOperationException("Could not connect to SharePoint Online.");
                }
                if (!string.IsNullOrEmpty(str))
                {
                    Version version;
                    string str2 = str.Split(new char[] { ',' })[0];
                    if (Version.TryParse(str2, out version))
                    {
                        return (version.Major >= 15);
                    }
                }
            }
            return false;
        }

        public static void ValidateUri(string name, Uri uri)
        {
            if ((string.IsNullOrEmpty(name) || (uri == null)) || ((uri.Scheme != Uri.UriSchemeHttp) && (uri.Scheme != Uri.UriSchemeHttps)))
            {
                throw new ArgumentException(name);
            }
        }

        public static void ValidateUrl(string name, string url)
        {
            Uri uri;
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentException(name);
            }
            try
            {
                uri = new Uri(url);
            }
            catch (UriFormatException)
            {
                throw new ArgumentException(name);
            }
            ValidateUri(name, uri);
        }

    }
}
