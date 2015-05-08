using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Lapointe.SharePointOnline.PowerShell.Data;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Set", "SPOList", SupportsShouldProcess = false, DefaultParameterSetName = "Default")]
    [CmdletGroup("Lists")]
    [CmdletDescription("Configures the specified list.", "The Set-SPOList cmdlet configures the list specified by the Identity parameter. Settings for any parameters that are not provided are not changed.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOList), typeof(NewSPOList))]
    [Example(Code = "PS C:\\> Set-SPOWeb -Identity \"/\" -CustomMasterUrl \"/_catalogs/masterpage/seattle.master\"",
        Remarks = "This example sets the master page of the root Site of a Site Collection to seattle.master.")]
    [Example(Code = "PS C:\\> Set-SPOWeb -Identity \"/childsite\" -BreakRoleInheritance -CopyRoleAssignments",
        Remarks = "This example breaks permissions inheritance for /childsite and copies the parents permissions to the Site.")]
    public class SetSPOList : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the List to update.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 1,
            ValueFromPipeline = true,
            Mandatory = false,
            HelpMessage = "Specifies the identity of the List to update.\r\n\r\nThe type must be a valid absolute URL, in the form http://contoso.sharepoint.com/site_name/lists/list_name, or an SPOList object, Microsoft.SharePoint.Client.List object, List title, or GUID representing the List ID.")]
        public SPOListPipeBind Identity { get; set; }

        [Parameter(
            ParameterSetName = "Default",
            Position = 1,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether Content Types are enabled for the List.")]
        public bool ContentTypesEnabled
        {
            get
            {
                return GetPropV<bool>("ContentTypesEnabled");
            }
            set
            {
                SetProp("ContentTypesEnabled", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 2,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the default workflow identifier for content approval on the List.")]
        public Guid DefaultContentApprovalWorkflowId
        {
            get
            {
                return GetPropV<Guid>("DefaultContentApprovalWorkflowId");
            }
            set
            {
                SetProp("DefaultContentApprovalWorkflowId", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 3,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the location of the default display form for the List.")]
        public string DefaultDisplayFormUrl
        {
            get
            {
                return GetProp<string>("DefaultDisplayFormUrl");
            }
            set
            {
                SetProp("DefaultDisplayFormUrl", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 4,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the URL of the edit form to use for List items in the List.")]
        public string DefaultEditFormUrl
        {
            get
            {
                return GetProp<string>("DefaultEditFormUrl");
            }
            set
            {
                SetProp("DefaultEditFormUrl", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 5,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the location of the default new form for the List.")]
        public string DefaultNewFormUrl
        {
            get
            {
                return GetProp<string>("DefaultNewFormUrl");
            }
            set
            {
                SetProp("DefaultNewFormUrl", value);
            }
        }

        [ValidateNotNull,
        Parameter(
            ParameterSetName = "Default",
            Position = 6,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the description of the List.")]
        public string Description
        {
            get
            {
                return GetProp<string>("Description");
            }
            set
            {
                SetProp("Description", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 7,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the reading order of the List. Valid values are NONE, LTR, or RTL.")]
        public string Direction
        {
            get
            {
                return GetProp<string>("Direction");
            }
            set
            {
                SetProp("Direction", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 8,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the server-relative URL of the document template for the List.")]
        public string DocumentTemplateUrl
        {
            get
            {
                return GetProp<string>("DocumentTemplateUrl");
            }
            set
            {
                SetProp("DocumentTemplateUrl", value);
            }
        }

        [ValidateNotNull,
        Parameter(
            ParameterSetName = "Default",
            Position = 9,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the minimum permission required to view minor versions and drafts within the List.")]
        public DraftVisibilityType DraftVersionVisibility
        {
            get
            {
                return GetPropV<DraftVisibilityType>("DraftVersionVisibility");
            }
            set
            {
                SetProp("DraftVersionVisibility", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 10,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether List item attachments are enabled for the List.")]
        public bool EnableAttachments
        {
            get
            {
                return GetPropV<bool>("EnableAttachments");
            }
            set
            {
                SetProp("EnableAttachments", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 11,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether new List folders can be added to the List.")]
        public bool EnableFolderCreation
        {
            get
            {
                return GetPropV<bool>("EnableFolderCreation");
            }
            set
            {
                SetProp("EnableFolderCreation", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 12,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether minor versions are enabled for the List.")]
        public bool EnableMinorVersions
        {
            get
            {
                return GetPropV<bool>("EnableMinorVersions");
            }
            set
            {
                SetProp("EnableMinorVersions", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 13,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether content approval is enabled for the List.")]
        public bool EnableModeration
        {
            get
            {
                return GetPropV<bool>("EnableModeration");
            }
            set
            {
                SetProp("EnableModeration", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 14,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether historical versions of List items and documents can be created in the List.")]
        public bool EnableVersioning
        {
            get
            {
                return GetPropV<bool>("EnableVersioning");
            }
            set
            {
                SetProp("EnableVersioning", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 15,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether List items must be checked out before they can be edited.")]
        public bool ForceCheckout
        {
            get
            {
                return GetPropV<bool>("ForceCheckout");
            }
            set
            {
                SetProp("ForceCheckout", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 16,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the List is hidden.")]
        public bool Hidden
        {
            get
            {
                return GetPropV<bool>("Hidden");
            }
            set
            {
                SetProp("Hidden", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 17,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the URI for the icon of the List.")]
        public string ImageUrl
        {
            get
            {
                return GetProp<string>("ImageUrl");
            }
            set 
            { 
                SetProp("ImageUrl", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 18,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether information rights management is enabled for the List.")]
        public bool IrmEnabled
        {
            get
            {
                return GetPropV<bool>("IrmEnabled");
            }
            set
            {
                SetProp("IrmEnabled", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 19,
            Mandatory = false)]
        public bool IrmExpire
        {
            get
            {
                return GetPropV<bool>("IrmExpire");
            }
            set
            {
                SetProp("IrmExpire", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 20,
            Mandatory = false)]
        public bool IrmReject
        {
            get
            {
                return GetPropV<bool>("IrmReject");
            }
            set
            {
                SetProp("IrmReject", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 21,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies a flag that a client application can use to determine whether to display the List.")]
        public bool IsApplicationList
        {
            get
            {
                return GetPropV<bool>("IsApplicationList");
            }
            set
            {
                SetProp("IsApplicationList", value);
            }
        }

        [ValidateNotNull,
        Parameter(
            ParameterSetName = "Default",
            Position = 22,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the last time a List item, field, or property of the List was modified.")]
        public DateTime LastItemModifiedDate
        {
            get
            {
                return GetPropV<DateTime>("LastItemModifiedDate");
            }
            set
            {
                SetProp("LastItemModifiedDate", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 23,
            Mandatory = false)]
        public bool MultipleDataList
        {
            get
            {
                return GetPropV<bool>("MultipleDataList");
            }
            set
            {
                SetProp("MultipleDataList", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 24,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies that the crawler must not crawl the List.")]
        public bool NoCrawl
        {
            get
            {
                return GetPropV<bool>("NoCrawl");
            }
            set
            {
                SetProp("NoCrawl", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 25,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether the List appears on the Quick Launch of the Site.")]
        public bool OnQuickLaunch
        {
            get
            {
                return GetPropV<bool>("OnQuickLaunch");
            }
            set
            {
                SetProp("OnQuickLaunch", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 26,
            Mandatory = false,
            HelpMessage = "Sets the displayed title for the List.")]
        public string Title
        {
            get
            {
                return GetProp<string>("Title");
            }
            set
            {
                SetProp("Title", value);
            }
        }

        [ValidateNotNull,
        Parameter(
            ParameterSetName = "Default",
            Position = 27,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the data validation criteria for a List item.")]
        public string ValidationFormula
        {
            get
            {
                return GetProp<string>("ValidationFormula");
            }
            set
            {
                SetProp("ValidationFormula", value);
            }
        }

        [ValidateNotNull,
        Parameter(
            ParameterSetName = "Default",
            Position = 28,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies the error message returned when data validation fails for a List item.")]
        public string ValidationMessage
        {
            get
            {
                return GetProp<string>("ValidationMessage");
            }
            set
            {
                SetProp("ValidationMessage", value);
            }
        }






        [Parameter(
            ParameterSetName = "BreakRoleInheritance",
            Position = 1,
            Mandatory = true,
            HelpMessage = "Creates unique role assignments for the Site.")]
        public SwitchParameter BreakRoleInheritance
        {
            get
            {
                return GetSwitch("BreakRoleInheritance");
            }
            set
            {
                SetProp("BreakRoleInheritance", value);
            }
        }

        [Parameter(
            ParameterSetName = "BreakRoleInheritance",
            Position = 2,
            Mandatory = false,
            HelpMessage = "Specifies whether to copy the role assignments from the parent Site.")]
        public SwitchParameter CopyRoleAssignments
        {
            get
            {
                return GetSwitch("CopyRoleAssignments");
            }
            set
            {
                SetProp("CopyRoleAssignments", value);
            }
        }

        [Parameter(
            ParameterSetName = "BreakRoleInheritance",
            Position = 3,
            Mandatory = false,
            HelpMessage = "The role assignments for all child Sites in the current Site and in the Sites which inherit role assignments from the current Site must be cleared and those Sites will inherit role assignments from the current Site after this call.")]
        public SwitchParameter ClearScopes
        {
            get
            {
                return GetSwitch("ClearScopes");
            }
            set
            {
                SetProp("ClearScopes", value);
            }
        }

        [Parameter(
            ParameterSetName = "ResetRoleInheritance",
            Position = 1,
            Mandatory = true,
            HelpMessage = "Resets the role inheritance for the Site and inherits role assignments from the parent Site.")]
        public SwitchParameter ResetRoleInheritance
        {
            get
            {
                return GetSwitch("ResetRoleInheritance");
            }
            set
            {
                SetProp("ResetRoleInheritance", value);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;
            var web = ctx.Site.OpenWeb(Web.Read());
            SPOList list = Identity.Read(web, true);
            
            if (base.ParameterSetName == "BreakRoleInheritance")
            {
                if (list.HasUniqueRoleAssignments)
                    throw new ArgumentException("The specified List already has unique role assignments.");

                if (base.Fields.Contains("BreakRoleInheritance") && BreakRoleInheritance)
                {
                    list.List.BreakRoleInheritance(CopyRoleAssignments, ClearScopes);
                    list.List.Context.ExecuteQuery();

                    Identity.Refresh();
                    list.Refresh();
                    WriteObject(list);
                    return;
                }
            }
            else if (base.ParameterSetName == "ResetRoleInheritance")
            {
                if (!list.HasUniqueRoleAssignments)
                    throw new ArgumentException("The specified List does not have unique role assignments.");

                if (base.Fields.Contains("ResetRoleInheritance") && ResetRoleInheritance)
                {
                    list.List.ResetRoleInheritance();
                    list.List.Context.ExecuteQuery();

                    Identity.Refresh();
                    list.Refresh();
                    WriteObject(list);
                    return;
                }
            }
            else
            {

                if (base.Fields.Contains("ContentTypesEnabled"))
                    list.ContentTypesEnabled = ContentTypesEnabled;

                if (base.Fields.Contains("DefaultContentApprovalWorkflowId"))
                    list.DefaultContentApprovalWorkflowId = DefaultContentApprovalWorkflowId;

                if (base.Fields.Contains("DefaultDisplayFormUrl"))
                    list.DefaultDisplayFormUrl = DefaultDisplayFormUrl;

                if (base.Fields.Contains("DefaultEditFormUrl"))
                    list.DefaultEditFormUrl = DefaultEditFormUrl;

                if (base.Fields.Contains("DefaultNewFormUrl"))
                    list.DefaultNewFormUrl = DefaultNewFormUrl;

                if (base.Fields.Contains("Description"))
                    list.Description = Description;

                if (base.Fields.Contains("Direction"))
                    list.Direction = Direction;

                if (base.Fields.Contains("DocumentTemplateUrl"))
                    list.DocumentTemplateUrl = DocumentTemplateUrl;

                if (base.Fields.Contains("DraftVersionVisibility"))
                    list.DraftVersionVisibility = DraftVersionVisibility;

                if (base.Fields.Contains("EnableAttachments"))
                    list.EnableAttachments = EnableAttachments;

                if (base.Fields.Contains("EnableFolderCreation"))
                    list.EnableFolderCreation = EnableFolderCreation;

                if (base.Fields.Contains("EnableMinorVersions"))
                    list.EnableMinorVersions = EnableMinorVersions;

                if (base.Fields.Contains("EnableModeration"))
                    list.EnableModeration = EnableModeration;

                if (base.Fields.Contains("EnableVersioning"))
                    list.EnableVersioning = EnableVersioning;

                if (base.Fields.Contains("ForceCheckout"))
                    list.ForceCheckout = ForceCheckout;

                if (base.Fields.Contains("Hidden"))
                    list.Hidden = Hidden;

                if (base.Fields.Contains("ImageUrl"))
                    list.ImageUrl = ImageUrl;

                if (base.Fields.Contains("IrmEnabled"))
                    list.IrmEnabled = IrmEnabled;

                if (base.Fields.Contains("IrmExpire"))
                    list.IrmExpire = IrmExpire;

                if (base.Fields.Contains("IrmReject"))
                    list.IrmReject = IrmReject;

                if (base.Fields.Contains("IsApplicationList"))
                    list.IsApplicationList = IsApplicationList;

                if (base.Fields.Contains("LastItemModifiedDate"))
                    list.LastItemModifiedDate = LastItemModifiedDate;

                if (base.Fields.Contains("MultipleDataList"))
                    list.MultipleDataList = MultipleDataList;

                if (base.Fields.Contains("NoCrawl"))
                    list.NoCrawl = NoCrawl;

                if (base.Fields.Contains("OnQuickLaunch"))
                    list.OnQuickLaunch = OnQuickLaunch;

                if (base.Fields.Contains("Title"))
                    list.Title = Title;

                if (base.Fields.Contains("ValidationFormula"))
                    list.ValidationFormula = ValidationFormula;

                if (base.Fields.Contains("ValidationMessage"))
                    list.ValidationMessage = ValidationMessage;

                list.Update();

                Identity.Refresh();
                list.Refresh();
                WriteObject(list);
            }
        }
    }
}
