using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Fields;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;

namespace Lapointe.SharePointOnline.PowerShell.Data.ContentTypes
{
    public class SPOContentType : SPOClientObject
    {
        private ContentType _ct = null;
        private bool _hasChanges = false;

        public SPOContentType(ContentType ct)
        {
            _ct = ct;
        }

        public ContentType ContentType { get { return _ct; } }

        public string Description
        {
            get
            {
                if (_ct.IsPropertyAvailable("Description"))
                    return _ct.Description;
                return null;
            }
            set
            {
                _ct.Description = value;
                _hasChanges = true;
            }
        }

        public string DisplayFormTemplateName
        {
            get
            {
                if (_ct.IsPropertyAvailable("DisplayFormTemplateName"))
                    return _ct.DisplayFormTemplateName;
                return null;
            }
            set
            {
                _ct.DisplayFormTemplateName = value;
                _hasChanges = true;
            }
        }

        public string DisplayFormUrl
        {
            get
            {
                if (_ct.IsPropertyAvailable("DisplayFormUrl"))
                    return _ct.DisplayFormUrl;
                return null;
            }
            set
            {
                _ct.DisplayFormUrl = value;
                _hasChanges = true;
            }
        }

        public string DocumentTemplate
        {
            get
            {
                if (_ct.IsPropertyAvailable("DocumentTemplate"))
                    return _ct.DocumentTemplate;
                return null;
            }
            set
            {
                _ct.DocumentTemplate = value;
                _hasChanges = true;
            }
        }

        public string DocumentTemplateUrl
        {
            get
            {
                if (_ct.IsPropertyAvailable("DocumentTemplateUrl"))
                    return _ct.DocumentTemplateUrl;
                return null;
            }
        }

        public string EditFormTemplateName
        {
            get
            {
                if (_ct.IsPropertyAvailable("EditFormTemplateName"))
                    return _ct.EditFormTemplateName;
                return null;
            }
            set
            {
                _ct.EditFormTemplateName = value;
                _hasChanges = true;
            }
        }

        public string EditFormUrl
        {
            get
            {
                if (_ct.IsPropertyAvailable("EditFormUrl"))
                    return _ct.EditFormUrl;
                return null;
            }
            set
            {
                _ct.EditFormUrl = value;
                _hasChanges = true;
            }
        }

        public FieldLinkCollection FieldLinks
        {
            get
            {
                if (_ct.IsObjectPropertyInstantiated("FieldLinks"))
                    return _ct.FieldLinks;
                return null;
            }
        }

        public FieldCollection Fields
        {
            get
            {
                if (_ct.IsObjectPropertyInstantiated("Fields"))
                    return _ct.Fields;
                return null;
            }
        }

        public string Group
        {
            get
            {
                if (_ct.IsPropertyAvailable("Group"))
                    return _ct.Group;
                return null;
            }
            set
            {
                _ct.Group = value;
                _hasChanges = true;
            }
        }

        public bool Hidden
        {
            get
            {
                if (_ct.IsPropertyAvailable("Hidden"))
                    return _ct.Hidden;
                return false;
            }
            set
            {
                _ct.Hidden = value;
                _hasChanges = true;
            }
        }

        public ContentTypeId Id
        {
            get
            {
                if (_ct.IsPropertyAvailable("Id"))
                    return _ct.Id;
                return null;
            }
        }

        public string JSLink
        {
            get
            {
                if (_ct.IsPropertyAvailable("JSLink"))
                    return _ct.JSLink;
                return null;
            }
            set
            {
                _ct.JSLink = value;
                _hasChanges = true;
            }
        }

        public string Name
        {
            get
            {
                if (_ct.IsPropertyAvailable("Name"))
                    return _ct.Name;
                return null;
            }
            set
            {
                _ct.Name = value;
                _hasChanges = true;
            }
        }

        public string NewFormTemplateName
        {
            get
            {
                if (_ct.IsPropertyAvailable("NewFormTemplateName"))
                    return _ct.NewFormTemplateName;
                return null;
            }
            set
            {
                _ct.NewFormTemplateName = value;
                _hasChanges = true;
            }
        }

        public string NewFormUrl
        {
            get
            {
                if (_ct.IsPropertyAvailable("NewFormUrl"))
                    return _ct.NewFormUrl;
                return null;
            }
            set
            {
                _ct.NewFormUrl = value;
                _hasChanges = true;
            }
        }

        public SPOContentType Parent
        {
            get
            {
                if (_ct.IsObjectPropertyInstantiated("Parent"))
                    return new SPOContentType(_ct.Parent);
                return null;
            }
        }

        public bool ReadOnly
        {
            get
            {
                if (_ct.IsPropertyAvailable("ReadOnly"))
                    return _ct.ReadOnly;
                return false;
            }
            set
            {
                _ct.ReadOnly = value;
                _hasChanges = true;
            }
        }

        public string SchemaXml
        {
            get
            {
                if (_ct.IsPropertyAvailable("SchemaXml"))
                    return _ct.SchemaXml;
                return null;
            }
        }

        public string SchemaXmlWithResourceTokens
        {
            get
            {
                if (_ct.IsPropertyAvailable("SchemaXmlWithResourceTokens"))
                    return _ct.SchemaXmlWithResourceTokens;
                return null;
            }
            set
            {
                _ct.SchemaXmlWithResourceTokens = value;
                _hasChanges = true;
            }
        }

        public string Scope
        {
            get
            {
                if (_ct.IsPropertyAvailable("Scope"))
                    return _ct.Scope;
                return null;
            }
        }

        public bool Sealed
        {
            get
            {
                if (_ct.IsPropertyAvailable("Sealed"))
                    return _ct.Sealed;
                return false;
            }
            set
            {
                _ct.Sealed = value;
                _hasChanges = true;
            }
        }

        public string StringId
        {
            get
            {
                if (_ct.IsPropertyAvailable("StringId"))
                    return _ct.StringId;
                return null;
            }
        }

        public WorkflowAssociationCollection WorkflowAssociations
        {
            get
            {
                if (_ct.IsObjectPropertyInstantiated("WorkflowAssociations"))
                    return _ct.WorkflowAssociations;
                return null;
            }
        }

        public void AddFieldLinks(SPOWeb web, params string[] fieldsToAdd)
        {
            AddFieldLinks(web.Web.AvailableFields, false, fieldsToAdd);
        }
        public void AddFieldLinks(SPOList list, params string[] fieldsToAdd)
        {
            AddFieldLinks(list.List.ParentWeb.AvailableFields, true, fieldsToAdd);
        }
        public void AddFieldLinks(Web web, params string[] fieldsToAdd)
        {
            AddFieldLinks(web.AvailableFields, false, fieldsToAdd);
        }
        public void AddFieldLinks(List list, params string[] fieldsToAdd)
        {
            AddFieldLinks(list.ParentWeb.AvailableFields, true, fieldsToAdd);
        }
        private void AddFieldLinks(FieldCollection fieldCollection, bool isList, params string[] fieldsToAdd)
        {
            List<SPOField> fields = new List<SPOField>();
            if (fieldsToAdd == null)
                return;
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            foreach (string fieldName in fieldsToAdd)
            {
                SPOField existingField = SPOField.GetField(ctx, fieldCollection, fieldName);
                if (existingField == null)
                    throw new ArgumentOutOfRangeException("Unable to locate field " + fieldName + ".");
                else
                    fields.Add(existingField);
            }
            if (fields.Count == 0)
                return;

            foreach (SPOField field in fields)
            {
                FieldLinkCreationInformation flci = new FieldLinkCreationInformation();
                flci.Field = field.Field;
                this.ContentType.FieldLinks.Add(flci);
            }
            if (fields.Count > 0)
            {
                this.ContentType.Update(!isList);
                ctx.ExecuteQuery();
            }
        }

        internal static SPOContentType GetContentType(CmdletContext ctx, ContentTypeCollection contentTypes, string name)
        {
            ctx.Load(contentTypes);
            ctx.ExecuteQuery();

            foreach (ContentType ct in contentTypes)
            {
                if (ct.Name.ToLower() == name.ToLower())
                {
                    SPOContentType.LoadContentType(ctx, ct);
                    return new SPOContentType(ct);
                }
            }
            return null;
        }

        public void Delete()
        {
            _ct.DeleteObject();
            _ct.Context.ExecuteQuery();
        }
        
        public void Refresh()
        {
            _ct.RefreshLoad();
            _ct.Context.ExecuteQuery();
        }

        public void Update(bool updateChildren)
        {
            if (_hasChanges)
            {
                _ct.Update(updateChildren);
                _ct.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }


        internal static void LoadContentType(CmdletContext ctx, ContentType ct)
        {
            ctx.Load(ct);
            ctx.Load(ct, 
                c => c.FieldLinks,
                c => c.Fields,
                c => c.Parent,
                c => c.SchemaXmlWithResourceTokens,
                c => c.WorkflowAssociations);

            ctx.ExecuteQuery();
        }
    }
}
