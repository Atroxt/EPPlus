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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Validation for decimal function results
    /// </summary>
    internal class DecimalCompileResultValidator : CompileResultValidator
    {
        /// <summary>
        /// Validate that decimal is not NaN or infinity
        /// </summary>
        /// <param name="obj"></param>
        /// <exception cref="ExcelErrorValueException"></exception>
        public override void Validate(object obj)
        {
            var num = ConvertUtil.GetValueDouble(obj);
            if (double.IsNaN(num) || double.IsInfinity(num))
            {
                throw new ExcelErrorValueException(eErrorType.Num);
            }
        }
    }
}
