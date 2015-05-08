using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.Lists
{
    public class SPOView : SPOClientObject
    {
        private View _view = null;
        private bool _hasChanges = false;
        private ViewCollection _parentCollection = null;

        public SPOView(View view, ViewCollection parentCollection)
        {
            _view = view;
            _parentCollection = parentCollection;
        }

        public View View { get { return _view; } }


        public string Aggregations
        {
            get
            {
                if (_view.IsPropertyAvailable("Aggregations"))
                    return _view.Aggregations;
                return null;
            }
            set
            {
                _view.Aggregations = value;
                _hasChanges = true;
            }
        }

        public string AggregationsStatus
        {
            get
            {
                if (_view.IsPropertyAvailable("AggregationsStatus"))
                    return _view.AggregationsStatus;
                return null;
            }
            set
            {
                _view.AggregationsStatus = value;
                _hasChanges = true;
            }
        }

        public string BaseViewId
        {
            get
            {
                if (_view.IsPropertyAvailable("BaseViewId"))
                    return _view.BaseViewId;
                return null;
            }
        }

        public ContentTypeId ContentTypeId
        {
            get
            {
                if (_view.IsPropertyAvailable("ContentTypeId"))
                    return _view.ContentTypeId;
                return null;
            }
            set
            {
                _view.ContentTypeId = value;
                _hasChanges = true;
            }
        }

        public bool DefaultView
        {
            get
            {
                if (_view.IsPropertyAvailable("DefaultView"))
                    return _view.DefaultView;
                return false;
            }
            set
            {
                _view.DefaultView = value;
                _hasChanges = true;
            }
        }

        public bool DefaultViewForContentType
        {
            get
            {
                if (_view.IsPropertyAvailable("DefaultViewForContentType"))
                    return _view.DefaultViewForContentType;
                return false;
            }
            set
            {
                _view.DefaultViewForContentType = value;
                _hasChanges = true;
            }
        }

        public bool EditorModified
        {
            get
            {
                if (_view.IsPropertyAvailable("EditorModified"))
                    return _view.EditorModified;
                return false;
            }
            set
            {
                _view.EditorModified = value;
                _hasChanges = true;
            }
        }

        public string Formats
        {
            get
            {
                if (_view.IsPropertyAvailable("Formats"))
                    return _view.Formats;
                return null;
            }
            set
            {
                _view.Formats = value;
                _hasChanges = true;
            }
        }

        public bool Hidden
        {
            get
            {
                if (_view.IsPropertyAvailable("Hidden"))
                    return _view.Hidden;
                return false;
            }
            set
            {
                _view.Hidden = value;
                _hasChanges = true;
            }
        }

        public string HtmlSchemaXml
        {
            get
            {
                if (_view.IsPropertyAvailable("HtmlSchemaXml"))
                    return _view.HtmlSchemaXml;
                return null;
            }
        }

        public Guid Id
        {
            get
            {
                if (_view.IsPropertyAvailable("Id"))
                    return _view.Id;
                return Guid.Empty;
            }
        }

        public string ImageUrl
        {
            get
            {
                if (_view.IsPropertyAvailable("ImageUrl"))
                    return _view.ImageUrl;
                return null;
            }
        }

        public bool IncludeRootFolder
        {
            get
            {
                if (_view.IsPropertyAvailable("IncludeRootFolder"))
                    return _view.IncludeRootFolder;
                return false;
            }
            set
            {
                _view.IncludeRootFolder = value;
                _hasChanges = true;
            }
        }

        public string JSLink
        {
            get
            {
                if (_view.IsPropertyAvailable("JSLink"))
                    return _view.JSLink;
                return null;
            }
            set
            {
                _view.JSLink = value;
                _hasChanges = true;
            }
        }

        public string ListViewXml
        {
            get
            {
                if (_view.IsPropertyAvailable("ListViewXml"))
                    return _view.ListViewXml;
                return null;
            }
            set
            {
                _view.ListViewXml = value;
                _hasChanges = true;
            }
        }

        public string Method
        {
            get
            {
                if (_view.IsPropertyAvailable("Method"))
                    return _view.Method;
                return null;
            }
            set
            {
                _view.Method = value;
                _hasChanges = true;
            }
        }

        public bool MobileDefaultView
        {
            get
            {
                if (_view.IsPropertyAvailable("MobileDefaultView"))
                    return _view.MobileDefaultView;
                return false;
            }
            set
            {
                _view.MobileDefaultView = value;
                _hasChanges = true;
            }
        }

        public bool MobileView
        {
            get
            {
                if (_view.IsPropertyAvailable("MobileView"))
                    return _view.MobileView;
                return false;
            }
            set
            {
                _view.MobileView = value;
                _hasChanges = true;
            }
        }

        public string ModerationType
        {
            get
            {
                if (_view.IsPropertyAvailable("ModerationType"))
                    return _view.ModerationType;
                return null;
            }
        }

        public bool OrderedView
        {
            get
            {
                if (_view.IsPropertyAvailable("OrderedView"))
                    return _view.OrderedView;
                return false;
            }
        }

        public bool Paged
        {
            get
            {
                if (_view.IsPropertyAvailable("Paged"))
                    return _view.Paged;
                return false;
            }
            set
            {
                _view.Paged = value;
                _hasChanges = true;
            }
        }

        public bool PersonalView
        {
            get
            {
                if (_view.IsPropertyAvailable("PersonalView"))
                    return _view.PersonalView;
                return false;
            }
        }

        public bool ReadOnlyView
        {
            get
            {
                if (_view.IsPropertyAvailable("ReadOnlyView"))
                    return _view.ReadOnlyView;
                return false;
            }
        }

        public bool RequiresClientIntegration
        {
            get
            {
                if (_view.IsPropertyAvailable("RequiresClientIntegration"))
                    return _view.RequiresClientIntegration;
                return false;
            }
        }

        public uint RowLimit
        {
            get
            {
                if (_view.IsPropertyAvailable("RowLimit"))
                    return _view.RowLimit;
                return 0;
            }
            set
            {
                _view.RowLimit = value;
                _hasChanges = true;
            }
        }

        public ViewScope Scope
        {
            get
            {
                if (_view.IsPropertyAvailable("Scope"))
                    return _view.Scope;
                return ViewScope.DefaultValue;
            }
            set
            {
                _view.Scope = value;
                _hasChanges = true;
            }
        }

        public string ServerRelativeUrl
        {
            get
            {
                if (_view.IsPropertyAvailable("ServerRelativeUrl"))
                    return _view.ServerRelativeUrl;
                return null;
            }
        }

        public string StyleId
        {
            get
            {
                if (_view.IsPropertyAvailable("StyleId"))
                    return _view.StyleId;
                return null;
            }
        }

        public bool Threaded
        {
            get
            {
                if (_view.IsPropertyAvailable("Threaded"))
                    return _view.Threaded;
                return false;
            }
        }

        public string Title
        {
            get
            {
                if (_view.IsPropertyAvailable("Title"))
                    return _view.Title;
                return null;
            }
            set
            {
                _view.Title = value;
                _hasChanges = true;
            }
        }

        public string Toolbar
        {
            get
            {
                if (_view.IsPropertyAvailable("Toolbar"))
                    return _view.Toolbar;
                return null;
            }
            set
            {
                _view.Toolbar = value;
                _hasChanges = true;
            }
        }

        public string ToolbarTemplateName
        {
            get
            {
                if (_view.IsPropertyAvailable("ToolbarTemplateName"))
                    return _view.ToolbarTemplateName;
                return null;
            }
        }

        public string ViewData
        {
            get
            {
                if (_view.IsPropertyAvailable("ViewData"))
                    return _view.ViewData;
                return null;
            }
            set
            {
                _view.ViewData = value;
                _hasChanges = true;
            }
        }

        public ViewFieldCollection ViewFields
        {
            get
            {
                if (_view.IsObjectPropertyInstantiated("ViewFields"))
                    return _view.ViewFields;
                return null;
            }
        }

        public string ViewJoins
        {
            get
            {
                if (_view.IsPropertyAvailable("ViewJoins"))
                    return _view.ViewJoins;
                return null;
            }
            set
            {
                _view.ViewJoins = value;
                _hasChanges = true;
            }
        }

        public string ViewProjectedFields
        {
            get
            {
                if (_view.IsPropertyAvailable("ViewProjectedFields"))
                    return _view.ViewProjectedFields;
                return null;
            }
            set
            {
                _view.ViewProjectedFields = value;
                _hasChanges = true;
            }
        }

        public string ViewQuery
        {
            get
            {
                if (_view.IsPropertyAvailable("ViewQuery"))
                    return _view.ViewQuery;
                return null;
            }
            set
            {
                _view.ViewQuery = value;
                _hasChanges = true;
            }
        }

        public string ViewType
        {
            get
            {
                if (_view.IsPropertyAvailable("ViewType"))
                    return _view.ViewType;
                return null;
            }
        }

        public void Update()
        {
            if (_hasChanges)
            {
                _view.Update();
                _view.Context.ExecuteQuery();
                _hasChanges = false;
            }
        }


        internal static void LoadView(CmdletContext ctx, View view)
        {
            ctx.Load(view);
            ctx.Load(view, v => v.ViewFields);

            ctx.ExecuteQuery();
        }

    }
}
