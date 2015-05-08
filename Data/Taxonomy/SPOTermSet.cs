using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Data.Taxonomy
{
    public class SPOTermSet : SPOTermSetItem
    {
        private TermSet _termSet = null;

        public SPOTermSet(TermSet termSet) : base(termSet)
        {
            _termSet = termSet;

            if (_termSet.IsPropertyAvailable("Description"))
                SetProp("Description", _termSet.Description, false);
            if (_termSet.IsPropertyAvailable("Contact"))
                SetProp("Contact", _termSet.Contact, false);
            if (_termSet.IsPropertyAvailable("IsOpenForTermCreation"))
                SetProp("IsOpenForTermCreation", _termSet.IsOpenForTermCreation, false);
        }

        public TermSet TermSet { get { return _termSet; } }

        public string Contact
        {
            get
            {
                return GetProp<string>("Contact");
            }
            set
            {
                SetProp("Contact", value, true);
            }
        }

        public string Description
        {
            get
            {
                return GetProp<string>("Description");
            }
            set
            {
                SetProp("Description", value, true);
            }
        }

        public bool IsOpenForTermCreation
        {
            get
            {
                return GetPropV<bool>("IsOpenForTermCreation");
            }
            set
            {
                SetProp("IsOpenForTermCreation", value, true);
            }
        }

        private SPOTermGroup _termGroup = null;
        public SPOTermGroup GetGroup()
        {
            if (_termGroup != null) return _termGroup;

            TermGroup group = _termSet.Group;
            _termSet.Context.Load(group);
            _termSet.Context.ExecuteQuery();
            _termGroup = new SPOTermGroup(group);
            return _termGroup;
        }

        public IEnumerable<string> Stakeholders
        {
            get { return _termSet.Stakeholders; }
        }

        public void AddStakeholder(string stakeholder)
        {
            _termSet.AddStakeholder(stakeholder);
            _termSet.Context.ExecuteQuery();
        }
        public void DeleteStakeholder(string stakeholder)
        {
            _termSet.DeleteStakeholder(stakeholder);
            _termSet.Context.ExecuteQuery();
        }

        public override void Refresh()
        {
            _termSet.RefreshLoad();
            _termSet.Context.ExecuteQuery();
            base.Refresh();

            if (_termSet.IsPropertyAvailable("Description"))
                SetProp("Description", _termSet.Description, false);
            if (_termSet.IsPropertyAvailable("Contact"))
                SetProp("Contact", _termSet.Contact, false);
            if (_termSet.IsPropertyAvailable("IsOpenForTermCreation"))
                SetProp("IsOpenForTermCreation", _termSet.IsOpenForTermCreation, false);
        }

        public override void Update()
        {
            if (HasChanges())
            {
                base.Update();

                if (HasChange("Contact"))
                    _termSet.Contact = Contact;

                if (HasChange("Description"))
                    _termSet.Description = Description;

                if (HasChange("IsOpenForTermCreation"))
                    _termSet.IsOpenForTermCreation = IsOpenForTermCreation;

                _termSet.Context.ExecuteQuery();
                Refresh();
            }
        }


    }
}
