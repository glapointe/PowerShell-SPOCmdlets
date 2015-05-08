using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data
{
    public abstract class SPOClientObject
    {
        private SPOClientObjectCollection _parentCollection;

        internal SPOClientObjectCollection ParentCollection
        {
            get
            {
                return _parentCollection;
            }
            set
            {
                _parentCollection = value;
            }
        }
    }
}
