using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lapointe.SharePointOnline.PowerShell.Data.Fields;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.PipeBind
{
    public sealed class SPOFieldPipeBind
    {
        private Guid _fieldId;
        private string _fieldNameOrTitle;
        private Field _field = null;

        public SPOFieldPipeBind(SPOField spoField)
        {
            if ((spoField == null) || spoField.Id == Guid.Empty)
            {
                throw new ArgumentNullException("The Field ID must be specified.");
            }
            _fieldId = spoField.Id;
            _field = spoField.Field;
        }
        public SPOFieldPipeBind(Field field)
        {
            if ((field == null) || field.Id == Guid.Empty)
            {
                throw new ArgumentNullException("The Field ID must be specified.");
            }
            _fieldId = field.Id;
            _field = field;
        }
        public SPOFieldPipeBind(Guid fieldId)
        {
            _fieldId = fieldId;
        }

        public SPOFieldPipeBind(string fieldNameOrTitle)
        {
            _fieldNameOrTitle = fieldNameOrTitle;
        }

        public string FieldNameOrTitle
        {
            get { return _fieldNameOrTitle; }
        }

        public Guid FieldId
        {
            get { return _fieldId; }
        }

        public SPOField Read(Web web)
        {
            return Read(web.AvailableFields);
        }
        public SPOField Read(List list)
        {
            return Read(list.Fields);
        }

        private SPOField Read(FieldCollection fields)
        {
            var ctx = SPOSiteContext.CurrentSiteContext.Context;
            Field field = null;
            if (this.FieldId != Guid.Empty)
            {
                field = fields.GetById(this.FieldId);
            }
            else if (!string.IsNullOrEmpty(this.FieldNameOrTitle))
            {
                field = fields.GetByInternalNameOrTitle(this.FieldNameOrTitle);
            }
            if (field != null)
            {
                _field = field;
                ctx.Load(field);
                ctx.ExecuteQuery();

                if (field.ServerObjectIsNull.Value)
                    return null;

                return new SPOField(_field);
            }
            return null;
        }
        public override string ToString()
        {
            if (this.FieldId != Guid.Empty)
                return this.FieldId.ToString();
            else if (!string.IsNullOrEmpty(this.FieldNameOrTitle))
                return this.FieldNameOrTitle;
            return "Field";
        }
        public void Refresh()
        {
            if (_field != null)
                new SPOField(_field).Refresh();
        }
    }
}
