using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Data
{
    public abstract class SPOClientObjectCollection : SPOClientObject, IEnumerable
    {
        private List<object> m_data;

        protected SPOClientObjectCollection()
        {
        }

        protected void AddChild(SPOClientObject obj)
        {
            this.Data.Add(obj);
            if (obj.ParentCollection == null)
            {
                obj.ParentCollection = this;
            }
        }

        protected object GetItemAtIndex(int i)
        {
            return this.Data[i];
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            int count = this.Data.Count;
            int i = 0;
            while (true)
            {
                if (i >= this.Data.Count)
                {
                    yield break;
                }
                if (count != this.Data.Count)
                {
                    throw new InvalidOperationException(Resources.GetString("CollectionModified"));
                }
                yield return this.GetItemAtIndex(i);
                i++;
            }
        }


        public virtual int Count
        {
            get
            {
                return this.Data.Count;
            }
        }

        protected List<object> Data
        {
            get
            {
                if (this.m_data == null)
                {
                    this.m_data = new List<object>();
                }
                return this.m_data;
            }
        }
    }


    public abstract class SPOClientObjectCollection<T> : SPOClientObjectCollection, IEnumerable<T>, IEnumerable
    {
        public SPOClientObjectCollection()
        {
        }

        public IEnumerator<T> GetEnumerator()
        {
            int count = this.Count;
            int i = 0;
            while (true)
            {
                if (i >= this.Count)
                {
                    yield break;
                }
                if (count != this.Count)
                {
                    throw new InvalidOperationException(Resources.GetString("CollectionModified"));
                }
                yield return (T) this.GetItemAtIndex(i);
                i++;
            }
        }

        public Type ElementType
        {
            get
            {
                return typeof(T);
            }
        }

        public System.Linq.Expressions.Expression Expression
        {
            get
            {
                return System.Linq.Expressions.Expression.Constant(this);
            }
        }

        public T this[int index]
        {
            get
            {
                return (T) base.GetItemAtIndex(index);
            }
        }

    }
}

