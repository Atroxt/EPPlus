﻿/*************************************************************************************************
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
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// 
    /// </summary>
    public partial class ExcelPivotTableFieldItemsCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>
    {
        ExcelPivotTableField _field;
        private Lookup<int, int> _cacheLookup = null;

        List<int> _hiddenItemIndex=null;
        internal ExcelPivotTableFieldItemsCollection(ExcelPivotTableField field) : base()
        {
            _field = field;            
        }
        internal void InitNewCalculation()
        {
            _hiddenItemIndex = null;
        }
        internal List<int> HiddenItemIndex
        {
            get
            {
                if (_hiddenItemIndex == null)
                {
                    _hiddenItemIndex = GetHiddenList();
                }
                return _hiddenItemIndex;
            }
        }

        private List<int> GetHiddenList()
        {
            List<int> hiddenItems = new List<int>();
            for (int i = 0; i < _list.Count; i++)
            {
                if (_list[i].Hidden)
                {
                    hiddenItems.Add(_list[i].X);
                }
            }
            return hiddenItems;
        }
        /// <summary>
        /// It the object exists in the cache
        /// </summary>
        /// <param name="value">The object to check for existance</param>
        /// <returns></returns>
        public bool Contains(object value)
        {
			var cl = _field.Cache.GetCacheLookup();
			return cl.ContainsKey(value);
        }
        /// <summary>
        /// Get the item with the value supplied. If the value does not exist, null is returned.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The pivot table field</returns>
        public ExcelPivotTableFieldItem GetByValue(object value)
        {
            if (value == null)
            {
                value = ExcelPivotTable.PivotNullValue;
            }

            var cl = _field.Cache.GetCacheLookup();
            if (cl.TryGetValue(value, out int ix))
            {
                if (CacheLookup.Contains(ix))
                {
                    return _list[CacheLookup[ix].First()];
                }
            }
			return null;
        }
        /// <summary>
        /// Get the index of the item with the value supplied. If the value does not exist, -1 is returned.
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns>The index of the item</returns>
        public int GetIndexByValue(object value)
        {
            if (value == null) return -1; 
            var cl = _field.Cache.GetCacheLookup();
			if (cl.TryGetValue(value, out int ix))
            {
                if (CacheLookup.Contains(ix))
                {
                    return _cacheLookup[ix].First();
                }
            }
            return -1;
        }
        internal void MatchValueToIndex()
        {
            var cache = _field.Cache;
            var isGroup = cache.Grouping != null;
            var cacheLookup = cache.GetCacheLookup();
            foreach (var item in _list)
            {
                var v = item.Value ?? ExcelPivotTable.PivotNullValue;
                if (item.Type == eItemType.Data && cacheLookup.TryGetValue(v, out int x))
                {
                    item.X = cacheLookup[v];
                }                
                else
                {
                    item.X = -1;
                }
            }
            _cacheLookup = null;
        }
        internal Lookup<int, int> CacheLookup
        {
            get
            {
                if (_cacheLookup == null)
                {
                    _cacheLookup = (Lookup<int, int>)_list.Where(x => x.X >= 0).ToLookup(x => x.X, y => _list.IndexOf(y));
                }
                return _cacheLookup;
            }
        }
        /// <summary>
        /// Set Hidden to false for all items in the collection
        /// </summary>
        public void ShowAll()
        {
            foreach(var item in _list)
            {
                item.Hidden = false;
            }
            _field.PageFieldSettings.SelectedItem = -1;
        }
        /// <summary>
        /// Set the ShowDetails for all items.
        /// </summary>
        /// <param name="isExpanded">The value of true is set all items to be expanded. The value of false set all items to be collapsed</param>
        public void ShowDetails(bool isExpanded=true)
        {
            if(!(_field.IsRowField || _field.IsColumnField))
            {
                //TODO: Add exception
            }
            if (_list.Count <= 1) Refresh();
            foreach (var item in _list)
            {
                item.ShowDetails= isExpanded;
            }
        }
        /// <summary>
        /// Hide all items except the item at the supplied index
        /// </summary>
        public void SelectSingleItem(int index)
        {
            if(index <0 || index >= _list.Count)
            {
                throw new ArgumentOutOfRangeException("index", "Index is out of range");
            }

            foreach (var item in _list)
            {
                if (item.Type == eItemType.Data)
                {
                    item.Hidden = true;
                }
            }
            _list[index].Hidden=false;
            if(_field.IsPageField)
            {
                _field.PageFieldSettings.SelectedItem = index;
            }
        }
        /// <summary>
        /// Refreshes the data of the cache field
        /// </summary>
        public void Refresh()
        {
            _field.Cache.Refresh();
            _hiddenItemIndex = null;
        }

		internal void Sort(eSortType sort)
		{
            var comparer = new PivotItemComparer(sort, _field);
			_list.Sort(comparer);
            _cacheLookup = null;
		}

        internal ExcelPivotTableFieldItem GetByCacheIndex(int index)
        {
            if (CacheLookup.Contains(index))
            {
                return _list[_cacheLookup[index].First()];
            }

            return null;
        }
	}
}