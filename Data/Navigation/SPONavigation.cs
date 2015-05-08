using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data.Navigation
{
    public class SPONavigation : SPOClientObject
    {
        Microsoft.SharePoint.Client.Navigation _navigation = null;
        SPONavigationNodeCollection _quickLaunch = null;
        SPONavigationNodeCollection _topNavigationBar = null;
        public SPONavigation(Microsoft.SharePoint.Client.Navigation navigation)
        {
            _navigation = navigation;
        }
        public Microsoft.SharePoint.Client.Navigation Navigation { get { return _navigation; } }

        public SPONavigationNodeCollection QuickLaunch
        {
            get
            {
                if (_quickLaunch != null)
                    return _quickLaunch;
                if (_navigation.IsObjectPropertyInstantiated("QuickLaunch"))
                {
                    _quickLaunch = new SPONavigationNodeCollection(_navigation.QuickLaunch);
                    return _quickLaunch;
                }
                return null;
            }
        }

        public SPONavigationNodeCollection TopNavigationBar
        {
            get
            {
                if (_topNavigationBar != null)
                    return _topNavigationBar;
                if (_navigation.IsObjectPropertyInstantiated("TopNavigationBar"))
                {
                    _topNavigationBar = new SPONavigationNodeCollection(_navigation.TopNavigationBar);
                    return _topNavigationBar;
                }
                return null;
            }
        }
        public bool UseShared { get; set; }

        public SPONavigationNode GetNodeById(int id)
        {
            return new SPONavigationNode(_navigation.GetNodeById(id));
        }


        internal static void LoadNavigation(CmdletContext ctx, Microsoft.SharePoint.Client.Navigation navigation)
        {
            ctx.Load(navigation, n => n.QuickLaunch, n => n.TopNavigationBar, n => n.UseShared);
            ctx.ExecuteQuery();
        }
    }
}
