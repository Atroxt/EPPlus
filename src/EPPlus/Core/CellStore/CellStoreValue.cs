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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections;

namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// This class represents
    /// </summary>
    internal class CellStoreValue : CellStore<ExcelValue> 
    {
        public CellStoreValue() : base()
        {

        }
        internal void SetValueRange_Value(int row, int col, object[,] array)
        {
            var rowBound = array.GetUpperBound(0);
            var colBound = array.GetUpperBound(1);
            
            for (int r = 0; r <= rowBound; r++)            
            {
                for (int c = 0; c <= colBound; c++)
                {
                    SetValue_Value(row + r, col + c, array[r, c]);
                }
            }
        }

        internal void SetValue_Value(int Row, int Column, object value)
        {
            var c = GetColumnIndex(Column);
            if(c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _value = value, _styleId = c._values[i]._styleId };
                    return;
                }
            }
            var v = new ExcelValue { _value = value, _styleId = GetStyleIdFromRowCol(Row, Column) };
            SetValue(Row, Column, v);
        }

        private int GetStyleIdFromRowCol(int row, int column)
        {
             var s = 0;
            if(row > 0)
            {
                s=GetValue(row, 0)._styleId;
            }
            if(s == 0 && column>0)
            {
                if (Exists(0, column))
                {
                    s= GetValue(0, column)._styleId;
                }
                else
                {
                    var r = 0;
                    var cp = GetColumnPosition(column);
                    if(GetPrevCell(ref r, ref cp, 0, 0, ColumnCount - 1))
                    {
                        var i=_columnIndex[cp].GetPointer(r);
                        if (i >= 0)
                        {
                            var prevCol = _columnIndex[cp]._values[i]._value as ExcelColumn;
                            if(prevCol!=null && prevCol.ColumnMax>=column)
                            {
                                s = prevCol.StyleID;
                            }
                        }
                    }

                }
            }
            return s;
        }

        internal void SetValue_Style(int Row, int Column, int styleId)
        {
            var c = GetColumnIndex(Column);
            if (c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _styleId = styleId, _value = c._values[i]._value };
                    return;
                }
            }
            var v = new ExcelValue { _styleId = styleId };
            SetValue(Row, Column, v);
        }
        internal void SetValue(int Row, int Column, object value, int styleId)
        {
            var c = GetColumnIndex(Column);
            if (c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _value = value, _styleId = styleId };
                    return;
                }
            }
            var v = new ExcelValue { _value = value, _styleId = styleId};
            SetValue(Row, Column, v);
        }

        internal int GetLastRow(int columnIndex)
        {
            if(columnIndex < ColumnCount)
            {
                var c = _columnIndex[columnIndex];
                if(c.PageCount>0)
                {
                    var p = c._pages[c.PageCount - 1];
                    return p.GetRow(p.RowCount-1);
                }
            }
            return 0;
        }

        internal int GetLastColumn()
        {
            if(ColumnCount>0 && _columnIndex[ColumnCount - 1].PageCount > 0)
            {
                var cIx = _columnIndex[ColumnCount - 1].GetPointer(0);
                if(cIx>=0)
                {
                    var c = _columnIndex[ColumnCount - 1]._values[cIx]._value as ExcelColumn;
                    if(c!=null)
                    {
                        return c.ColumnMax;
                    }
                }
            }
            return 0;
        }

    }
}