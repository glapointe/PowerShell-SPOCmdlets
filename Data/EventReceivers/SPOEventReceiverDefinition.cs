using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;

namespace Lapointe.SharePointOnline.PowerShell.Data.EventReceivers
{
    public class SPOEventReceiverDefinition : SPOClientObject
    {
        private EventReceiverDefinition _erd = null;
        private EventReceiverDefinitionCollection _parentCollection = null;

        public SPOEventReceiverDefinition(EventReceiverDefinition erd, EventReceiverDefinitionCollection parentCollection)
        {
            _erd = erd;
            _parentCollection = parentCollection;
        }

        public EventReceiverDefinition EventReceiverDefinition { get { return _erd; } }

        public EventReceiverType EventType
        {
            get
            {
                if (_erd.IsPropertyAvailable("EventType"))
                    return _erd.EventType;
                return EventReceiverType.ContextEvent;
            }
        }

        public string ReceiverAssembly
        {
            get
            {
                if (_erd.IsPropertyAvailable("ReceiverAssembly"))
                    return _erd.ReceiverAssembly;
                return null;
            }
        }

        public string ReceiverClass
        {
            get
            {
                if (_erd.IsPropertyAvailable("ReceiverClass"))
                    return _erd.ReceiverClass;
                return null;
            }
        }

        public Guid ReceiverId
        {
            get
            {
                if (_erd.IsPropertyAvailable("ReceiverId"))
                    return _erd.ReceiverId;
                return Guid.Empty;
            }
        }

        public string ReceiverName
        {
            get
            {
                if (_erd.IsPropertyAvailable("ReceiverName"))
                    return _erd.ReceiverName;
                return null;
            }
        }

        public string ReceiverUrl
        {
            get
            {
                if (_erd.IsPropertyAvailable("ReceiverUrl"))
                    return _erd.ReceiverUrl;
                return null;
            }
        }

        public int SequenceNumber
        {
            get
            {
                if (_erd.IsPropertyAvailable("SequenceNumber"))
                    return _erd.SequenceNumber;
                return 0;
            }
        }

        public EventReceiverSynchronization Synchronization
        {
            get
            {
                if (_erd.IsPropertyAvailable("Synchronization"))
                    return _erd.Synchronization;
                return EventReceiverSynchronization.DefaultSynchronization;
            }
        }

        public void Delete()
        {
            _erd.DeleteObject();
            _parentCollection.Context.ExecuteQuery();
        }
    }
}
