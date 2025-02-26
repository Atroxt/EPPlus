﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/7/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Sorting.Internal;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting
{
    internal class RangeSorter
    {
        public RangeSorter(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        private readonly ExcelWorksheet _worksheet;

        private void ValidateColumnArray(ExcelRangeBase range, int[] columns)
        {
            var cols = range._toCol - range._fromCol + 1;
            foreach (var c in columns)
            {
                if (c > cols - 1 || c < 0)
                {
                    throw (new ArgumentException("Cannot reference columns outside the boundries of the range. Note that column references are zero-based within the range"));
                }
            }
        }

        private void ValidateRowsArray(ExcelRangeBase range, int[] rows)
        {
            var nRows = range._toRow - range._fromRow + 1;
            foreach (var r in rows)
            {
                if (r > nRows - 1 || r < 0)
                {
                    throw (new ArgumentException("Cannot reference rows outside the boundries of the range. Note that row references are zero-based within the range"));
                }
            }
        }

        private bool[] CreateDefaultDescendingArray(int[] sortParams)
        {
            var descending = new bool[sortParams.Length];
            for (int i = 0; i < sortParams.Length; i++)
            {
                descending[i] = false;
            }
            return descending;
        }

        private ExcelRangeBase GetSortRange(ExcelRangeBase range, ExcelWorksheet ws)
        {
            if (ws.Dimension == null) return null;
            var dimension = ws.Dimension;
            var fromRow = range._fromRow < dimension.Start.Row ? dimension.Start.Row : range._fromRow;
            var toRow = range._toRow > dimension.End.Row ? dimension.End.Row : range._toRow;
            return ws.Cells[fromRow, range._fromCol, toRow, range._toCol];
        }

        private void ClearRowsAfter(ExcelRangeBase range, int nRows)
        {
            var startCol = range._fromCol;
            var endCol = range._toCol;
            var ws = range.Worksheet;
            for(var row = range.Start.Row + nRows; row <= range.End.Row; row++)
            {
                ws.Cells[row, startCol, row, endCol].Clear();
            }
        }

        public void Sort(
            ExcelRangeBase range, 
            int[] columns, 
            ref bool[] descending, 
            CultureInfo culture = null, 
            CompareOptions compareOptions = CompareOptions.None, 
            Dictionary<int, string[]> customLists = null)
        {
            if (columns == null)
            {
                columns = new int[] { 0 };
            }
            ValidateColumnArray(range, columns);
            if (descending == null)
            {
                descending = CreateDefaultDescendingArray(columns);
            }
            var ws = range.Worksheet;
            var r = GetSortRange(range, ws);
            if (r == null) return;
            var sortItems = SortItemFactory.Create(r);
            var comp = new EPPlusSortComparer(columns, descending, customLists, culture ?? CultureInfo.CurrentCulture, compareOptions);
            sortItems.Sort(comp);
            var wsd = new RangeWorksheetData(r);

            ApplySortedRange(r, sortItems, wsd);
            ClearRowsAfter(r, sortItems.Count);
        }

        public void SortLeftToRight(
            ExcelRangeBase range,
            int[] rows,
            ref bool[] descending,
            CultureInfo culture,
            CompareOptions compareOptions = CompareOptions.None,
            Dictionary<int, string[]> customLists = null
            )
        {
            if (rows == null)
            {
                rows = new int[] { 0 };
            }
            ValidateRowsArray(range, rows);
            if (descending == null)
            {
                descending = CreateDefaultDescendingArray(rows);
            }
            var sortItems = SortItemLeftToRightFactory.Create(range);
            var comp = new EPPlusSortComparerLeftToRight(rows, descending, customLists, culture ?? CultureInfo.CurrentCulture, compareOptions);
            sortItems.Sort(comp);
            var wsd = new RangeWorksheetData(range);

            ApplySortedRange(range, sortItems, wsd);
        }

        private void ApplySortedRange(ExcelRangeBase range, List<SortItem<ExcelValue>> sortItems, RangeWorksheetData wsd)
        {
            //Sort the values and styles.
            var nColumnsInRange = range._toCol - range._fromCol + 1;
            var dim = range.Worksheet.Dimension;
            if(dim != null && nColumnsInRange > dim.End.Column)
            {
                nColumnsInRange = dim.End.Column;
            }
            _worksheet._values.Clear(range._fromRow, range._fromCol, range._toRow - range._fromRow + 1, nColumnsInRange);
            for (var r = 0; r < sortItems.Count; r++)
            {
                for (int c = 0; c < nColumnsInRange; c++)
                {
                    var row = range._fromRow + r;
                    var col = range._fromCol + c;
                    _worksheet._values.SetValue(row, col, sortItems[r].Items[c]);
                    var addr = ExcelCellBase.GetAddress(sortItems[r].Row, range._fromCol + c);
                    //Move flags
                    HandleFlags(wsd, row, col, addr);
                    //Move metadata
                    HandleMetadata(wsd, row, col, addr);

                    //Move formulas
                    HandleFormula(wsd, row, col, addr, sortItems[r].Row, col);

                    //Move hyperlinks
                    HandleHyperlink(wsd, row, col, addr);

                    //Move comments
                    HandleComment(wsd, row, col, addr);

                    //Move threaded comments
                    HandleThreadedComment(wsd, row, col, addr);
                }
            }
            if(sortItems.Count < range.Rows)
            {
                var delFromRow = range._fromRow + sortItems.Count;
                //Clear comments in the store, otherwise the clear of the range might delete comments that have been moved in the sort operation.
                _worksheet._commentsStore.Delete(delFromRow, range._fromCol, range._toRow - delFromRow+1, nColumnsInRange, false);
                _worksheet._threadedCommentsStore.Delete(delFromRow, range._fromCol, range._toRow - delFromRow + 1, nColumnsInRange, false);
            }
        }

        private void ApplySortedRange(ExcelRangeBase range, List<SortItemLeftToRight<ExcelValue>> sortItems, RangeWorksheetData wsd)
        {
            //Sort the values and styles.
            var nRowsInRange = range._toRow - range._fromRow + 1;
            var dim = range.Worksheet.Dimension;
            if (dim != null && nRowsInRange > dim.End.Row)
            {
                nRowsInRange = dim.End.Row;
            }

            _worksheet._values.Clear(range._fromRow, range._fromCol, range._toRow - range._fromRow + 1, range._toCol);
            for (var c = 0; c < sortItems.Count; c++)
            {
                for (int r = 0; r < nRowsInRange; r++)
                {
                    var row = range._fromRow + r;
                    var col = range._fromCol + c;
                    //_worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
                    _worksheet._values.SetValue(row, col, sortItems[c].Items[r]);
                    var addr = ExcelCellBase.GetAddress(range._fromRow + r, sortItems[c].Column);
                    //Move flags
                    HandleFlags(wsd, row, col, addr);
                    //Move metadata
                    HandleMetadata(wsd, row, col, addr);

                    //Move formulas
                    HandleFormula(wsd, row, col, addr, row, sortItems[c].Column);

                    //Move hyperlinks
                    HandleHyperlink(wsd, row, col, addr);

                    //Move comments
                    HandleComment(wsd, row, col, addr);

                    //Move threaded comments
                    HandleThreadedComment(wsd, row, col, addr);
                }
            }
            if (sortItems.Count < range.Columns)
            {
                var delFromCol = range._fromCol + sortItems.Count;
                //Clear comments in the store, otherwise the clear of the range might delete comments that have been moved in the sort operation.
                _worksheet._commentsStore.Delete(range._fromRow, delFromCol, nRowsInRange, range._toCol - delFromCol + 1, false);
                _worksheet._threadedCommentsStore.Delete(range._fromRow, delFromCol, nRowsInRange, range._toCol - delFromCol + 1, false);
            }
        }

        private void HandleHyperlink(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Hyperlinks.ContainsKey(addr))
            {
                _worksheet._hyperLinks.SetValue(row, col, wsd.Hyperlinks[addr]);
            }
        }

        private void HandleMetadata(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Metadata.ContainsKey(addr))
            {
                _worksheet._metadataStore.SetValue(row, col, wsd.Metadata[addr]);
            }
            else
            {
                _worksheet._metadataStore.Clear(row, col, 1, 1);
            }
        }

        private void HandleFlags(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Flags.ContainsKey(addr))
            {
                _worksheet._flags.SetValue(row, col, wsd.Flags[addr]);
            }
            else
            {
                _worksheet._flags.Clear(row, col, 1, 1);
            }
        }

        private void HandleComment(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.Comments.ContainsKey(addr))
            {
                var i = wsd.Comments[addr];
                _worksheet._commentsStore.SetValue(row, col, i);
                var comment = _worksheet._comments._list[i];
                comment.Reference = ExcelCellBase.GetAddress(row, col);
            }
            else
            {
                _worksheet._commentsStore.Clear(row, col, 1, 1);
            }
        }

        private void HandleThreadedComment(RangeWorksheetData wsd, int row, int col, string addr)
        {
            if (wsd.ThreadedComments.ContainsKey(addr))
            {
                var i = wsd.ThreadedComments[addr];
                _worksheet._threadedCommentsStore.SetValue(row, col, i);
                var threadedComment = _worksheet._threadedComments._threads[i];
                threadedComment.SetAddress(ExcelCellBase.GetAddress(row, col));
            }
        }

        private void HandleFormula(RangeWorksheetData wsd, int row, int col, string addr, int initialRow, int initialCol)
        {
            if (wsd.Formulas.ContainsKey(addr))
            {
                _worksheet._formulas.SetValue(row, col, wsd.Formulas[addr]);
                if(wsd.Formulas[addr] is string)
                {
                    var formula = wsd.Formulas[addr].ToString();
                    var newFormula = initialRow != row ?
                        AddressUtility.ShiftAddressRowsInFormula(string.Empty, formula, 1, row - initialRow) :
                        AddressUtility.ShiftAddressColumnsInFormula(string.Empty, formula, 1, col - initialCol);
                    _worksheet._formulas.SetValue(row, col, newFormula);
                }
                else if (wsd.Formulas[addr] is int)
                {
                    int sfIx = (int)wsd.Formulas[addr];
                    var startAddr = new ExcelAddress(_worksheet._sharedFormulas[sfIx].Address);
                    var f = _worksheet._sharedFormulas[sfIx];

                    f.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), row, col);
                    f.Address = ExcelCellBase.GetAddress(row, col, row, col);
                }
            }
            else
            {
                _worksheet._formulas.Clear(row, col, 1, 1);
            }
        }

        internal void SetWorksheetSortState(ExcelRangeBase range, int[] columnsOrRows, bool[] descending, CompareOptions compareOptions, bool leftToRight, Dictionary<int, string[]> customLists)
        {
            //Set sort state
            var sortState = new SortState(_worksheet.NameSpaceManager, _worksheet);
            sortState.Ref = range.Address;
            sortState.ColumnSort = leftToRight;
            sortState.CaseSensitive = (compareOptions == CompareOptions.IgnoreCase || compareOptions == CompareOptions.OrdinalIgnoreCase);
            for (var ix = 0; ix < columnsOrRows.Length; ix++)
            {
                bool? desc = null;
                if (descending.Length > ix && descending[ix])
                {
                    desc = true;
                }
                var adr = leftToRight ?
                    ExcelCellBase.GetAddress(range._fromRow + columnsOrRows[ix], range._fromCol, range._fromRow + columnsOrRows[ix], range._toCol) :
                    ExcelCellBase.GetAddress(range._fromRow, range._fromCol + columnsOrRows[ix], range._toRow, range._fromCol + columnsOrRows[ix]);
                if (customLists != null && customLists.ContainsKey(columnsOrRows[ix]))
                {
                    sortState.SortConditions.Add(adr, desc, customLists[columnsOrRows[ix]]);
                }
                else
                {
                    sortState.SortConditions.Add(adr, desc);
                }
            }
        }
    }
}
