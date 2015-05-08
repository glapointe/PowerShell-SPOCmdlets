using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Linq.Expressions;
using System.Management.Automation;
using System.Text;
using System.Threading;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets
{
    public class BaseSPOCmdlet : PSCmdlet
    {
        private CmdletContext _context;
        private HybridDictionary _fields;

        public BaseSPOCmdlet()
        {
            _fields = new HybridDictionary();
        }

        protected CmdletContext Context
        {
            get
            {
                if (_context != null)
                {
                    return _context;
                }
                if (SPOSiteContext.CurrentSiteContext == null)
                {
                    throw new InvalidOperationException("No connection available. Use Connect-SPOSite before running this CmdLet.");
                }
                return SPOSiteContext.CurrentSiteContext.Context;
            }
            set
            {
                _context = value;
            }
        }

        protected T GetProp<T>(string fieldName) where T : class
        {
            return (T)this.Fields[fieldName];
        }


        protected S GetPropV<S>(string fieldName) where S : struct
        {
            if (!this.Fields.Contains(fieldName))
            {
                throw new InvalidCastException("uninitialized property cannot be unboxed.");
            }
            return (S)this.Fields[fieldName];
        }


        protected SwitchParameter GetSwitch(string fieldName)
        {
            if (!this.Fields.Contains(fieldName))
            {
                SwitchParameter parameter = new SwitchParameter(false);
                this.Fields[fieldName] = parameter;
            }
            return (SwitchParameter)this.Fields[fieldName];
        }

        protected void SetProp(string fieldName, object value)
        {
            if (value != null)
            {
                if (value.GetType() == typeof(string))
                {
                    value = ((string)value).Trim();
                }
                this.Fields[fieldName] = value;
            }
            else
            {
                this.Fields.Remove(fieldName);
            }
        }

 


        protected HybridDictionary Fields
        {
            get
            {
                return _fields;
            }
        }
 

 

 

    }
}
