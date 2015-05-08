using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public class SPOTermLabel : SPOTaxonomyBase
    {
        private Label _label = null;
        private SPOTerm _term = null;

        public SPOTermLabel(Label label, SPOTerm term)
        {
            _label = label;
            _term = term;

            if (_label.IsPropertyAvailable("Value"))
                SetProp("Value", _label.Value, false);
        }

        public Label Label { get { return _label; } }

        public string Value
        {
            get
            {
                return GetProp<string>("Value");
            }
            set
            {
                SetProp("Value", value, true);
            }
        }

        public bool IsDefaultForLanguage
        {
            get
            {
                if (_label.IsPropertyAvailable("IsDefaultForLanguage"))
                    return _label.IsDefaultForLanguage;
                return false;
            }
        }

        public int Language
        {
            get
            {
                if (_label.IsPropertyAvailable("Language"))
                    return _label.Language;
                return 0;
            }
        }

        public SPOTerm Term { get { return _term; } }

        public void Delete()
        {
            _label.DeleteObject();
            _label.Context.ExecuteQuery();
        }

        public void SetAsDefaultForLanguage()
        {
            _label.SetAsDefaultForLanguage();
            _label.Context.ExecuteQuery();
            Refresh();
        }

        public void Refresh()
        {
            _label.RefreshLoad();
            _label.Context.ExecuteQuery();

            if (_label.IsPropertyAvailable("Value"))
                SetProp("Value", _label.Value, false);
        }

        public void Update()
        {
            if (HasChanges())
            {
                if (HasChange("Value"))
                    _label.Value = Value;
                _label.Context.ExecuteQuery();
                Refresh();
            }
        }

    }
}
