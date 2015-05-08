using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.PipeBind;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("New", "SPOFile", SupportsShouldProcess = false, DefaultParameterSetName = "List")]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new list file within an existing List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList))]
    [Example(Code = "PS C:\\> New-SPOFile -Web \"/\" -List \"My List\" -File \"c:\\myfile.txt\" -FieldValues @{\"Title\"=\"My new file\"}",
        Remarks = "This example creates a new file within the List My List under the root Site of the current Site Collection.")]
    public class NewSPOFile : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to add the file to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "List",
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List to add the file to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Folder",
            Position = 2,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the Folder to add the file to.\r\n\r\nThe type must be a valid SPOFolder object.")]
        public SPOFolder Folder { get; set; }

        [ValidateNotNullOrEmpty,
        ValidateFileExists,
        Parameter(
            Position = 3,
            Mandatory = true,
            HelpMessage = "Specify the path to the file to add to the list.")]
        public string File { get; set; }

        [Parameter(Position = 4, HelpMessage = "Overwrite an existing file if present.")]
        public SwitchParameter Overwrite { get; set; }

        [Parameter(
            Position = 5,
            Mandatory = false,
            HelpMessage = "The collection of field values to set where the key is the internal field name. The type must be a hash table where each key represents the name of a field whose value should be set to the corresponding key value (e.g., @{\"Field1\"=\"Value1\";\"Field2\"=\"Value2\"}). Alternatively, provide the path to a file with XML property settings (<Properties><Property Name=\"Name1\">Value1</Property><Property Name=\"Name2\">Value2</Property></Properties>).")]
        public PropertiesPipeBind FieldValues { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            Dictionary<string, object> fieldValues = null;
            if (FieldValues != null)
                fieldValues = FieldValues.Read();

            byte[] content = System.IO.File.ReadAllBytes(File);
            System.IO.FileInfo file = new System.IO.FileInfo(File);
            string leafName = file.Name;

            if (ParameterSetName == "List")
            {
                var web = ctx.Site.OpenWeb(Web.Read());
                SPOList list = List.Read(web, true);
                WriteObject(list.AddFile(content, Overwrite, leafName, fieldValues));
            }
            else
            {
                WriteObject(Folder.AddFile(content, Overwrite, leafName, fieldValues));
            }
        }
    }
}
