using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.PipeBind;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("New", "SPOListItem", SupportsShouldProcess = false)]
    [CmdletGroup("Lists")]
    [CmdletDescription("Creates a new list item within an existing List.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList))]
    [Example(Code = "PS C:\\> New-SPOListItem -Web \"/\" -List \"My List\" -FieldValues @{\"Title\"=\"My new list item\"}",
        Remarks = "This example creates a new list item within the List My List under the root Site of the current Site Collection.")]
    public class NewSPOListItem : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to add the item to.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the List to add the item to.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind List { get; set; }

        [Parameter(
            Position = 2,
            Mandatory = false,
            HelpMessage = "Create the item within a subfolder inside the list. Only include the path relative to the root of the list.")]
        public string ListRelativeUrlFolderPath { get; set; }

        [Parameter(
            Position = 3,
            Mandatory = false,
            HelpMessage = "The collection of field values to set where the key is the internal field name. The type must be a hash table where each key represents the name of a field whose value should be set to the corresponding key value (e.g., @{\"Field1\"=\"Value1\";\"Field2\"=\"Value2\"}). Alternatively, provide the path to a file with XML property settings (<Properties><Property Name=\"Name1\">Value1</Property><Property Name=\"Name2\">Value2</Property></Properties>).")]
        public PropertiesPipeBind FieldValues { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            var web = ctx.Site.OpenWeb(Web.Read());
            SPOList list = List.Read(web, true);

            Dictionary<string, object> fieldValues = null;
            if (FieldValues != null)
                fieldValues = FieldValues.Read();

            WriteObject(list.AddItem(ListRelativeUrlFolderPath, fieldValues));
        }
    }
}
