using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Fields
{
    public class SPOField : SPOClientObject
    {
        private Field _field = null;
        private bool _hasChanges = false;

        public SPOField(Field field)
        {
            _field = field;
        }

        public Field Field { get { return _field; } }

        public bool CanBeDeleted
        {
            get
            {
                if (_field.IsPropertyAvailable("CanBeDeleted"))
                    return _field.CanBeDeleted;
                return false;
            }
        }
        public string DefaultValue
        {
            get
            {
                if (_field.IsPropertyAvailable("DefaultValue"))
                    return _field.DefaultValue;
                return null;
            }
            set
            {
                _field.DefaultValue = value;
                _hasChanges = true;
            }
        }
        public string Description
        {
            get
            {
                if (_field.IsPropertyAvailable("Description"))
                    return _field.Description;
                return null;
            }
            set
            {
                _field.Description = value;
                _hasChanges = true;
            }
        }
        public string Direction
        {
            get
            {
                if (_field.IsPropertyAvailable("Direction"))
                    return _field.Direction;
                return null;
            }
            set
            {
                _field.Direction = value;
                _hasChanges = true;
            }
        }
        public bool EnforceUniqueValues
        {
            get
            {
                if (_field.IsPropertyAvailable("EnforceUniqueValues"))
                    return _field.EnforceUniqueValues;
                return false;
            }
            set
            {
                _field.EnforceUniqueValues = value;
                _hasChanges = true;
            }
        }
        public string EntityPropertyName
        {
            get
            {
                if (_field.IsPropertyAvailable("EntityPropertyName"))
                    return _field.EntityPropertyName;
                return null;
            }
        }
        public FieldType FieldTypeKind
        {
            get
            {
                if (_field.IsPropertyAvailable("FieldTypeKind"))
                    return _field.FieldTypeKind;
                return FieldType.Invalid;
            }
            set
            {
                _field.FieldTypeKind = value;
                _hasChanges = true;
            }
        }
        public bool Filterable
        {
            get
            {
                if (_field.IsPropertyAvailable("Filterable"))
                    return _field.Filterable;
                return false;
            }
        }
        public bool FromBaseType
        {
            get
            {
                if (_field.IsPropertyAvailable("FromBaseType"))
                    return _field.FromBaseType;
                return false;
            }
        }
        public string Group
        {
            get
            {
                if (_field.IsPropertyAvailable("Group"))
                    return _field.Group;
                return null;
            }
            set
            {
                _field.Group = value;
                _hasChanges = true;
            }
        }
        public bool Hidden
        {
            get
            {
                if (_field.IsPropertyAvailable("Hidden"))
                    return _field.Hidden;
                return false;
            }
            set
            {
                _field.Hidden = value;
                _hasChanges = true;
            }
        }
        public Guid Id
        {
            get
            {
                if (_field.IsPropertyAvailable("Id"))
                    return _field.Id;
                return Guid.Empty;
            }
        }
        public bool Indexed
        {
            get
            {
                if (_field.IsPropertyAvailable("Indexed"))
                    return _field.Indexed;
                return false;
            }
            set
            {
                _field.Indexed = value;
                _hasChanges = true;
            }
        }
        public string InternalName
        {
            get
            {
                if (_field.IsPropertyAvailable("InternalName"))
                    return _field.InternalName;
                return null;
            }
        }
        public string JSLink
        {
            get
            {
                if (_field.IsPropertyAvailable("JSLink"))
                    return _field.JSLink;
                return null;
            }
            set
            {
                _field.JSLink = value;
                _hasChanges = true;
            }
        }
        public bool ReadOnlyField
        {
            get
            {
                if (_field.IsPropertyAvailable("ReadOnlyField"))
                    return _field.ReadOnlyField;
                return false;
            }
            set
            {
                _field.ReadOnlyField = value;
                _hasChanges = true;
            }
        }
        public bool Required
        {
            get
            {
                if (_field.IsPropertyAvailable("Required"))
                    return _field.Required;
                return false;
            }
            set
            {
                _field.Required = value;
                _hasChanges = true;
            }
        }
        public string SchemaXml
        {
            get
            {
                if (_field.IsPropertyAvailable("SchemaXml"))
                    return _field.SchemaXml;
                return null;
            }
            set
            {
                _field.SchemaXml = value;
                _hasChanges = true;
            }
        }
        public string SchemaXmlWithResourceTokens
        {
            get
            {
                if (_field.IsPropertyAvailable("SchemaXmlWithResourceTokens"))
                    return _field.SchemaXmlWithResourceTokens;
                return null;
            }
        }
        public string Scope
        {
            get
            {
                if (_field.IsPropertyAvailable("Scope"))
                    return _field.Scope;
                return null;
            }
        }
        public bool Sealed
        {
            get
            {
                if (_field.IsPropertyAvailable("Sealed"))
                    return _field.Sealed;
                return false;
            }
        }
        public bool Sortable
        {
            get
            {
                if (_field.IsPropertyAvailable("Sortable"))
                    return _field.Sortable;
                return false;
            }
        }
        public string StaticName
        {
            get
            {
                if (_field.IsPropertyAvailable("StaticName"))
                    return _field.StaticName;
                return null;
            }
            set
            {
                _field.StaticName = value;
                _hasChanges = true;
            }
        }
        public string Title
        {
            get
            {
                if (_field.IsPropertyAvailable("Title"))
                    return _field.Title;
                return null;
            }
            set
            {
                _field.Title = value;
                _hasChanges = true;
            }
        }
        public string TypeAsString
        {
            get
            {
                if (_field.IsPropertyAvailable("TypeAsString"))
                    return _field.TypeAsString;
                return null;
            }
            set
            {
                _field.TypeAsString = value;
                _hasChanges = true;
            }
        }
        public string TypeDisplayName
        {
            get
            {
                if (_field.IsPropertyAvailable("TypeDisplayName"))
                    return _field.TypeDisplayName;
                return null;
            }
        }
        public string TypeShortDescription
        {
            get
            {
                if (_field.IsPropertyAvailable("TypeShortDescription"))
                    return _field.TypeShortDescription;
                return null;
            }
        }

        public string ValidationFormula
        {
            get
            {
                if (_field.IsPropertyAvailable("ValidationFormula"))
                    return _field.ValidationFormula;
                return null;
            }
            set
            {
                _field.ValidationFormula = value;
                _hasChanges = true;
            }
        }
        public string ValidationMessage
        {
            get
            {
                if (_field.IsPropertyAvailable("ValidationMessage"))
                    return _field.ValidationMessage;
                return null;
            }
            set
            {
                _field.ValidationMessage = value;
                _hasChanges = true;
            }
        }

        public void SetFormDisplay(bool showInDisplayForm, bool showInEditForm, bool showInNewForm)
        {
            _field.SetShowInDisplayForm(showInDisplayForm);
            _field.SetShowInEditForm(showInEditForm);
            _field.SetShowInNewForm(showInNewForm);
            _field.Context.ExecuteQuery();
        }
        public void SetShowInDisplayForm(bool value)
        {
            _field.SetShowInDisplayForm(value);
            _field.Context.ExecuteQuery();
        }
        public void SetShowInEditForm(bool value)
        {
            _field.SetShowInEditForm(value);
            _field.Context.ExecuteQuery();
        }
        public void SetShowInNewForm(bool value)
        {
            _field.SetShowInNewForm(value);
            _field.Context.ExecuteQuery();
        }


        public void Delete()
        {
            _field.DeleteObject();
            _field.Context.ExecuteQuery();
        }


        internal static SPOField GetField(CmdletContext ctx, FieldCollection fields, XmlDocument schemaXmlDoc)
        {
            Guid fieldId = Guid.Parse(schemaXmlDoc.DocumentElement.GetAttribute("Id"));
            Field field = fields.GetById(fieldId);
            ctx.Load(field);
            ctx.ExecuteQuery();
            if (!field.ServerObjectIsNull.Value)
            {
                return new SPOField(field);
            }
            field = fields.GetByInternalNameOrTitle(schemaXmlDoc.DocumentElement.GetAttribute("Name"));
            ctx.Load(field);
            ctx.ExecuteQuery();
            if (!field.ServerObjectIsNull.Value)
            {
                return new SPOField(field);
            }
            return null;
        }
        internal static SPOField GetField(CmdletContext ctx, FieldCollection fields, string nameOrTitle)
        {
            Field field = fields.GetByInternalNameOrTitle(nameOrTitle);
            ctx.Load(field);
            ctx.ExecuteQuery();
            if (!field.ServerObjectIsNull.Value)
            {
                return new SPOField(field);
            }
            return null;
        }

        public void Refresh()
        {
            _field.RefreshLoad();
            _field.Context.ExecuteQuery();
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _field.Update();
                _field.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }

        public void UpdateAndPushChanges(bool pushChangesToLists)
        {
            if (_hasChanges)
            {
                _field.UpdateAndPushChanges(pushChangesToLists);
                _field.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }
    }
}
