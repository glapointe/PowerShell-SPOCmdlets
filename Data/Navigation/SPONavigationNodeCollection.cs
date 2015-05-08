using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data.Navigation
{
    public class SPONavigationNodeCollection : SPOClientObjectCollection<SPONavigationNode>
    {
        Microsoft.SharePoint.Client.NavigationNodeCollection _navigationNodeCollection = null;

        public SPONavigationNodeCollection(Microsoft.SharePoint.Client.NavigationNodeCollection navigationNodeCollection)
        {
            _navigationNodeCollection = navigationNodeCollection;
            foreach (var node in _navigationNodeCollection)
            {
                AddChild(new SPONavigationNode(node));
            }
        }

        public Microsoft.SharePoint.Client.NavigationNodeCollection NavigationNodeCollection { get { return _navigationNodeCollection; } }

        public SPONavigationNode Add(Microsoft.SharePoint.Client.NavigationNodeCreationInformation parameters)
        {
            var node = new SPONavigationNode(_navigationNodeCollection.Add(parameters));
            _navigationNodeCollection.Context.Load(node.NavigationNode);
            _navigationNodeCollection.Context.ExecuteQuery();
            AddChild(node);
            return node;
        }
        public SPONavigationNode Add(bool asLastNode, bool isExternal, SPONavigationNode previousNode, string title, string url)
        {
            var node = new Microsoft.SharePoint.Client.NavigationNodeCreationInformation();
            node.AsLastNode = asLastNode;
            node.IsExternal = isExternal;
            if (previousNode != null)
                node.PreviousNode = previousNode.NavigationNode;
            node.Title = title;
            node.Url = url;
            return Add(node);
        }

        internal void Remove(SPONavigationNode node)
        {
            Data.Remove(node);
        }


    }
}
