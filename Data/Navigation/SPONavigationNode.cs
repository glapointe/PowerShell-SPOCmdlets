using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data.Navigation
{
    public class SPONavigationNode : SPOClientObject
    {
        Microsoft.SharePoint.Client.NavigationNode _navigationNode = null;
        private bool _hasChanges = false;
        SPONavigationNodeCollection _children = null;

        public SPONavigationNode(Microsoft.SharePoint.Client.NavigationNode navigationNode)
        {
            _navigationNode = navigationNode;
        }

        public Microsoft.SharePoint.Client.NavigationNode NavigationNode { get { return _navigationNode; } }

        public SPONavigationNodeCollection GetChildren()
        {
            if (_children != null)
                return _children;
            if (!_navigationNode.IsObjectPropertyInstantiated("Children"))
            {
                _navigationNode.Context.Load(_navigationNode.Children);
                _navigationNode.Context.ExecuteQuery();
            }
            _children = new SPONavigationNodeCollection(_navigationNode.Children);
            return _children;
        }

        public int Id
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("Id"))
                    return _navigationNode.Id;
                return 0;
            }
        }
        public bool IsDocLib
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("IsDocLib"))
                    return _navigationNode.IsDocLib;
                return false;
            }
        }

        public bool IsExternal
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("IsExternal"))
                    return _navigationNode.IsExternal;
                return false;
            }
        }
        public bool IsVisible
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("IsVisible"))
                    return _navigationNode.IsVisible;
                return false;
            }
            set
            {
                _navigationNode.IsVisible = value;
                _hasChanges = true;
            }
        }


        public string Title
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("Title"))
                    return _navigationNode.Title;
                return null;
            }
            set
            {
                _navigationNode.Title = value;
                _hasChanges = true;
            }
        }

        public string Url
        {
            get
            {
                if (_navigationNode.IsPropertyAvailable("Url"))
                    return _navigationNode.Url;
                return null;
            }
            set
            {
                _navigationNode.Url = value;
                _hasChanges = true;
            }
        }



        public void Delete()
        {
            _navigationNode.DeleteObject();
            _navigationNode.Context.ExecuteQuery();
            if (ParentCollection != null && ParentCollection is SPONavigationNodeCollection)
                ((SPONavigationNodeCollection)ParentCollection).Remove(this);
        }
        public void Update()
        {
            if (_hasChanges)
            {
                _navigationNode.Update();
                _navigationNode.Context.Load(_navigationNode);
                _navigationNode.Context.ExecuteQuery();

                _hasChanges = false;
            }
        }
    }
}
