using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;

namespace Lapointe.SharePointOnline.PowerShell.Attributes
{
    
    public class ValidateDirectoryExistsAttribute : ValidateArgumentsAttribute
    {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics)
        {
            string str = arguments as string;
            if (string.IsNullOrEmpty(str))
            {
                throw new PSArgumentNullException();
            }

            if (!Directory.Exists(str))
            {
                throw new PSArgumentException("Directory not found");
            }
        }
    }
}
