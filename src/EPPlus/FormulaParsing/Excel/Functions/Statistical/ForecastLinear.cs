﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "6.0",
        Description = "Calculate, or predict, a future value by using existing values. The future value is a y-value for a given x-value.")]
    internal class ForecastLinear : ExcelFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var x = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var arg1 = arguments[1];
            var arg2 = arguments[2];
            var arrayY = ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg1 }, context).Select(a => a.Value).ToArray();
            var arrayX = ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg2 }, context).Select(b => b.Value).ToArray();
            if (arrayY.Count() != arrayX.Count()) return CompileResult.GetErrorResult(eErrorType.NA);
            if (!arrayY.Any()) return CompileResult.GetErrorResult(eErrorType.NA);
            var result = Forecast.ForecastImpl(x, arrayY, arrayX);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
