// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Logging;
using ApplicationCore.Interfaces;


namespace ApplicationCore.Helpers
{
    public abstract class CollectionPage<T> : ICollectionPage<T>
    {
        private string _skipToken = String.Empty;
        private int _itemsPage = 10;
        private int _pageIndex = 1;

        /// <summary>
        /// Default constructor
        /// </summary>
        public CollectionPage()
        {
            this.CurrentPage = new List<T>();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="currentPage"></param>
        public CollectionPage(IList<T> currentPage)
        {
            this.CurrentPage = currentPage;
        }

        public string SkipToken { get { return _skipToken; } }

        public int ItemsPage { get { return _itemsPage; } }

        public int PageIndex { get { return _pageIndex; } }

        public IList<T> CurrentPage { get; private set; }

        #region IList methods

        public int IndexOf(T item)
        {
            return this.CurrentPage.IndexOf(item);
        }

        public void Insert(int index, T item)
        {
            this.CurrentPage.Insert(index, item);
        }

        public void RemoveAt(int index)
        {
            this.CurrentPage.RemoveAt(index);
        }

        public T this[int index]
        {
            get { return this.CurrentPage[index]; }
            set { this.CurrentPage[index] = value; }
        }

        public void Add(T item)
        {
            this.CurrentPage.Add(item);
        }

        public void Clear()
        {
            this.CurrentPage.Clear();
        }

        public bool Contains(T item)
        {
            return this.CurrentPage.Contains(item);
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            this.CurrentPage.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get { return this.CurrentPage.Count; }
        }

        public bool IsReadOnly
        {
            get { return this.CurrentPage.IsReadOnly; }
        }

        public bool Remove(T item)
        {
            return this.CurrentPage.Remove(item);
        }

        public IEnumerator<T> GetEnumerator()
        {
            return this.CurrentPage.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.CurrentPage.GetEnumerator();
        }
        #endregion

        // Private methods
    }
}
