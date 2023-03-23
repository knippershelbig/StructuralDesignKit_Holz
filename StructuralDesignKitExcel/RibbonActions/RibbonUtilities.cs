using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StructuralDesignKitExcel.RibbonActions
{
    internal class RibbonUtilities
    {
        /// <summary>
        /// Create a cell validation with a given list
        /// </summary>
        /// <param name="cell">cell to modify</param>
        /// <param name="list">List to implement</param>
        public static void ValidateCellWithList(Range cell, List<string> list)
        {
            if (cell == null) { throw new Exception("Please open a new workbook first"); }
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

            string separator = ",";
            if (System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator == ";") separator = ";";

            var flatList = string.Join(separator, list.ToArray());
            string initialValue = list[0];


            cell.Validation.Delete();
            cell.Validation.Add(
            XlDVType.xlValidateList,
            XlDVAlertStyle.xlValidAlertInformation,
            XlFormatConditionOperator.xlBetween,
            flatList,
            Type.Missing);
            cell.Validation.IgnoreBlank = true;
            cell.Validation.InCellDropdown = true;
            cell.Value2 = initialValue;
        }
    }
}
