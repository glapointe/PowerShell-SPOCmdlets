using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public abstract class SPOTaxonomyBase : SPOClientObject
    {
        protected Dictionary<string, object> _original = new Dictionary<string, object>();
        protected Dictionary<string, object> _changes = new Dictionary<string, object>();

        protected T GetProp<T>(string fieldName) where T : class
        {
            if (_changes.ContainsKey(fieldName))
                return (T)_changes[fieldName];
            if (_original.ContainsKey(fieldName))
                return (T)_original[fieldName];

            return null;
        }


        protected S GetPropV<S>(string fieldName) where S : struct
        {
            if (!_changes.ContainsKey(fieldName) && !_original.ContainsKey(fieldName))
            {
                throw new InvalidCastException("Uninitialized property cannot be unboxed.");
            }

            if (_changes.ContainsKey(fieldName))
                return (S)_changes[fieldName];
            else
                return (S)_original[fieldName];
        }

        protected void SetProp(string fieldName, object value, bool isChange)
        {
            Dictionary<string, object> collection = null;
            if (isChange)
                collection = _changes;
            else
                collection = _original;

            if (collection.ContainsKey(fieldName))
                collection[fieldName] = value;
            else
                collection.Add(fieldName, value);

            if (!isChange && _changes.ContainsKey(fieldName))
                _changes.Remove(fieldName);
        }

        protected bool HasChange(string fieldName)
        {
            return _changes.ContainsKey(fieldName);
        }
        protected bool HasChanges()
        {
            return _changes.Count > 0;
        }
        protected void ClearChanges()
        {
            _changes.Clear();
        }

        protected int _workingLanguage = 0;
        protected int GetWorkingLanguage(TermStore termStore)
        {
            if (_workingLanguage > 0) return _workingLanguage;
            if (!termStore.IsPropertyAvailable("WorkingLanguage"))
            {
                termStore.Context.Load(termStore, l => l.WorkingLanguage);
                termStore.Context.ExecuteQuery();
            }
            _workingLanguage = termStore.WorkingLanguage;
            return _workingLanguage;
        }

    }
}
