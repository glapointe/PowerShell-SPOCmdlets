using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public abstract class CmdletPipeBind<CmdletObject> where CmdletObject: class
    {
        protected CmdletPipeBind()
        {
        }

        protected CmdletPipeBind(CmdletObject inputObject)
        {
            if (inputObject == null)
            {
                throw new ArgumentNullException("inputObject");
            }
            this.Discover(inputObject);
        }

        public static void CheckUrlSchema(string name, Uri uri)
        {
            if (uri == null)
            {
                throw new ArgumentNullException("uri");
            }
            if ((uri.Scheme != Uri.UriSchemeHttp) && (uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new ArgumentException(name);
            }
        }

        protected abstract void Discover(CmdletObject instance);
        public abstract CmdletObject Read();
        public static void ValidateNotNull(string name, object value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(name);
            }
        }

        public static void ValidateString(string name, string value)
        {
            CmdletPipeBind<CmdletObject>.ValidateNotNull(name, value);
            if (value.Trim().Length == 0)
            {
                throw new ArgumentException(name);
            }
        }

        public static void ValidateUri(string name, Uri uri)
        {
            CmdletPipeBind<CmdletObject>.CheckUrlSchema(name, uri);
        }

        public static void ValidateUrl(string name, string url)
        {
            Uri uri;
            CmdletPipeBind<CmdletObject>.ValidateString(name, url);
            try
            {
                uri = new Uri(url);
            }
            catch (UriFormatException)
            {
                throw new ArgumentException(name);
            }
            CmdletPipeBind<CmdletObject>.CheckUrlSchema(name, uri);
        }
    }
}
