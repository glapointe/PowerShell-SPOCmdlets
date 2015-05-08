using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Exceptions
{
    public class FileExistsException : Exception
    {
        public FileExistsException(string message) : base(message) { }
    }
}
