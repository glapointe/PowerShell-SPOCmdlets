using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public class UrlPipeBind : CmdletPipeBind<Uri>
    {
        private Uri m_url;

        public UrlPipeBind(string inputObject) : base(new Uri(inputObject))
        {
        }

        public UrlPipeBind(Uri inputObject) : base(inputObject)
        {
        }

        protected override void Discover(Uri instance)
        {
            this.m_url = instance;
        }

        public override Uri Read()
        {
            return this.m_url;
        }
    }
}
