/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Base collection class for pivottable fields
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
    {
        internal List<T> _list = new List<T>();
        internal ExcelPivotTableFieldCollectionBase()
        {
        }
        /// <summary>
        /// Gets the enumerator of the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        internal void AddInternal(T field)
        {
            _list.Add(field);
        }
        internal void Clear()
        {
            _list.Clear();
        }
        /// <summary>
        /// Indexer for the  collection
        /// </summary>
        /// <param name="Index">The index</param>
        /// <returns>The pivot table field</returns>
        public virtual T this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _list.Count)
                {
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }
                return _list[Index];
            }
        }
        /// <summary>
        /// Returns the zero-based index of the item.
        /// </summary>
        /// <param name="item">The item</param>
        /// <returns>the zero-based index of the item in the list</returns>
        internal int IndexOf(T item)
        {
            return _list.IndexOf(item);
        }
    }
}