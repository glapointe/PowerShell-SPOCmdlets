using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;

namespace Lapointe.SharePointOnline.PowerShell.Attributes
{
    
    public class ValidateDirectoryExistsAndValidFileNameAttribute : ValidateArgumentsAttribute
    {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics)
        {
            string str = arguments as string;
            if (string.IsNullOrEmpty(str))
            {
                throw new PSArgumentNullException();
            }

            FileInfo info = new FileInfo(str);
            if (info.Directory.Exists)
            {
                if (info.Name.EndsWith("\\") || info.Name.EndsWith("/"))
                {
                    throw new PSArgumentException("Filename not specified");
                }
            }
            else
            {
                throw new PSArgumentException("Directory not found");
            }
        }
    }
}
