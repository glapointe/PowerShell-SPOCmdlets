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

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Sites
{
    [Cmdlet("Set", "SPOWeb", SupportsShouldProcess = false, DefaultParameterSetName = "Default")]
    [CmdletGroup("Sites")]
    [CmdletDescription("Configures the specified subsite.", "The Set-SPOWeb cmdlet configures the subsite specified by the Identity parameter. Settings for any parameters that are not provided are not changed.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(GetSPOWeb), typeof(NewSPOWeb))]
    [Example(Code = "PS C:\\> Set-SPOWeb -Identity \"/\" -CustomMasterUrl \"/_catalogs/masterpage/seattle.master\"",
        Remarks = "This example sets the master page of the root Site of a Site Collection to seattle.master.")]
    [Example(Code = "PS C:\\> Set-SPOWeb -Identity \"/childsite\" -BreakRoleInheritance -CopyRoleAssignments",
        Remarks = "This example breaks permissions inheritance for /childsite and copies the parents permissions to the Site.")]
    public class SetSPOWeb : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site to update.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the SIte ID.")]
        public SPOWebPipeBind Identity { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 1,
            Mandatory = false,
            HelpMessage = "Sets the URL for a custom master page file to apply to the Site.")]
        public string CustomMasterUrl
        {
            get
            {
                return GetProp<string>("CustomMasterUrl");
            }
            set
            {
                SetProp("CustomMasterUrl", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 2,
            Mandatory = false,
            HelpMessage = "Sets the URL of the master page that is used for the Site.")]
        public string MasterUrl
        {
            get
            {
                return GetProp<string>("MasterUrl");
            }
            set
            {
                SetProp("MasterUrl", value);
            }
        }

        [ValidateNotNull,
        Alias("Name"),
        Parameter(
            ParameterSetName = "Default",
            Position = 3,
            Mandatory = false,
            HelpMessage = "Specifies the new title of the Site.")]
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
            Position = 4,
            Mandatory = false,
            HelpMessage = "Specifies the new description of the Site.")]
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

        [Parameter(
            ParameterSetName = "Default",
            Position = 5,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the Site should use Minimal Download Strategy.")]
        public bool EnableMinimalDownload
        {
            get
            {
                return GetPropV<bool>("EnableMinimalDownload");
            }
            set
            {
                SetProp("EnableMinimalDownload", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 6,
            Mandatory = false,
            HelpMessage = "Sets a value that specifies whether the Quick Launch area is enabled on the Site.")]
        public bool QuickLaunchEnabled
        {
            get
            {
                return GetPropV<bool>("QuickLaunchEnabled");
            }
            set
            {
                SetProp("QuickLaunchEnabled", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 7,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the Site can be saved as a Site template.")]
        public bool SaveSiteAsTemplateEnabled
        {
            get
            {
                return GetPropV<bool>("SaveSiteAsTemplateEnabled");
            }
            set
            {
                SetProp("SaveSiteAsTemplateEnabled", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Alias("RelativeUrl"),
        Parameter(
            ParameterSetName = "Default",
            Position = 8,
            Mandatory = false,
            HelpMessage = "Specifies the new relative URL for the Site. This is the URL path after the site collection URL.")]
        public string ServerRelativeUrl
        {
            get
            {
                return GetProp<string>("ServerRelativeUrl");
            }
            set
            {
                SetProp("ServerRelativeUrl", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 9,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the RSS feeds are enabled on the Site.")]
        public bool SyndicationEnabled
        {
            get
            {
                return GetPropV<bool>("SyndicationEnabled");
            }
            set
            {
                SetProp("SyndicationEnabled", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 10,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the tree view is enabled on the Site.")]
        public bool TreeViewEnabled
        {
            get
            {
                return GetPropV<bool>("TreeViewEnabled");
            }
            set
            {
                SetProp("TreeViewEnabled", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 11,
            Mandatory = false,
            HelpMessage = "Sets the user interface (UI) version of the Site.")]
        public int UIVersion
        {
            get
            {
                return GetPropV<int>("UIVersion");
            }
            set
            {
                SetProp("UIVersion", value);
            }
        }

        [Parameter(
            ParameterSetName = "Default",
            Position = 12,
            Mandatory = false,
            HelpMessage = "Sets a Boolean value that specifies whether the settings UI for visual upgrade is shown or hidden.")]
        public bool UIVersionConfigurationEnabled
        {
            get
            {
                return GetPropV<bool>("UIVersionConfigurationEnabled");
            }
            set
            {
                SetProp("UIVersionConfigurationEnabled", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 13,
            Mandatory = false,
            HelpMessage = "Sets the associated member group of the Site.")]
        public SPOGroupPipeBind AssociatedMemberGroup
        {
            get
            {
                return GetProp<SPOGroupPipeBind>("AssociatedMemberGroup");
            }
            set
            {
                SetProp("AssociatedMemberGroup", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 14,
            Mandatory = false,
            HelpMessage = "Sets the associated owner group of the Site.")]
        public SPOGroupPipeBind AssociatedOwnerGroup
        {
            get
            {
                return GetProp<SPOGroupPipeBind>("AssociatedOwnerGroup");
            }
            set
            {
                SetProp("AssociatedOwnerGroup", value);
            }
        }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "Default",
            Position = 15,
            Mandatory = false,
            HelpMessage = "Sets the associated visitor group of the Site.")]
        public SPOGroupPipeBind AssociatedVisitorGroup
        {
            get
            {
                return GetProp<SPOGroupPipeBind>("AssociatedVisitorGroup");
            }
            set
            {
                SetProp("AssociatedVisitorGroup", value);
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
            var cweb = ctx.Site.OpenWeb(Identity.Read());
            SPOWeb.LoadWeb(ctx, cweb, true);
            SPOWeb web = new SPOWeb(cweb);

            if (base.ParameterSetName == "BreakRoleInheritance")
            {
                if (web.HasUniqueRoleAssignments)
                    throw new ArgumentException("The specified Site already has unique role assignments.");

                if (base.Fields.Contains("BreakRoleInheritance") && BreakRoleInheritance)
                {
                    web.Web.BreakRoleInheritance(CopyRoleAssignments, ClearScopes);
                    web.Web.Context.ExecuteQuery();

                    Identity.Refresh();
                    web.Refresh();
                    WriteObject(web);
                    return;
                }
            }
            else if (base.ParameterSetName == "ResetRoleInheritance")
            {
                if (!web.HasUniqueRoleAssignments)
                    throw new ArgumentException("The specified Site does not have unique role assignments.");

                if (base.Fields.Contains("ResetRoleInheritance") && ResetRoleInheritance)
                {
                    web.Web.ResetRoleInheritance();
                    web.Web.Context.ExecuteQuery();

                    Identity.Refresh();
                    web.Refresh();
                    WriteObject(web);
                    return;
                }
            }
            else
            {
                if ((base.Fields.Contains("AssociatedMemberGroup") ||
                    base.Fields.Contains("AssociatedOwnerGroup") ||
                    base.Fields.Contains("AssociatedVisitorGroup")) && !web.HasUniqueRoleAssignments)
                {
                    throw new ArgumentException("You can only set the associated member, owner, and visitor groups when the site has unique permissions (does not inherit).");
                }

                if (base.Fields.Contains("CustomMasterUrl"))
                    web.CustomMasterUrl = CustomMasterUrl;

                if (base.Fields.Contains("MasterUrl"))
                    web.MasterUrl = MasterUrl;

                if (base.Fields.Contains("Title"))
                    web.Title = Title;

                if (base.Fields.Contains("Description"))
                    web.Description = Description;

                if (base.Fields.Contains("EnableMinimalDownload"))
                    web.EnableMinimalDownload = EnableMinimalDownload;

                if (base.Fields.Contains("QuickLaunchEnabled"))
                    web.QuickLaunchEnabled = QuickLaunchEnabled;

                if (base.Fields.Contains("SaveSiteAsTemplateEnabled"))
                    web.SaveSiteAsTemplateEnabled = SaveSiteAsTemplateEnabled;

                if (base.Fields.Contains("ServerRelativeUrl"))
                    web.ServerRelativeUrl = ServerRelativeUrl;

                if (base.Fields.Contains("SyndicationEnabled"))
                    web.SyndicationEnabled = SyndicationEnabled;

                if (base.Fields.Contains("TreeViewEnabled"))
                    web.TreeViewEnabled = TreeViewEnabled;

                if (base.Fields.Contains("UIVersion"))
                    web.UIVersion = UIVersion;

                if (base.Fields.Contains("UIVersionConfigurationEnabled"))
                    web.UIVersionConfigurationEnabled = UIVersionConfigurationEnabled;

                if (base.Fields.Contains("AssociatedMemberGroup"))
                    web.AssociatedMemberGroup = AssociatedMemberGroup.Read();

                if (base.Fields.Contains("AssociatedOwnerGroup"))
                    web.AssociatedOwnerGroup = AssociatedOwnerGroup.Read();

                if (base.Fields.Contains("AssociatedVisitorGroup"))
                    web.AssociatedVisitorGroup = AssociatedVisitorGroup.Read();

                web.Update();

                Identity.Refresh();
                web.Refresh();
                WriteObject(web);
            }
        }
    }
}
