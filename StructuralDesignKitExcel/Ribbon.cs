using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary;
using StructuralDesignKitLibrary.Materials;

namespace StructuralDesignKitExcel
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @" 
    <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='SDK' label='SDK Timber'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
              <button id='button2' label='2nd Fumction' onAction='OnFormulaFunction' tag='SDK.Materials.Fmk'/>
              <button id='CrossSectionTag' label='Creates a SDK Cross section Tag' onAction='CreatesCrossSection'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;


            var cell = xlApp.ActiveCell;

            ValidateCellWithMaterials(cell);    
           
        }

        public void OnFormulaFunction(IRibbonControl control)
        {
            //myfunction.TestFunction()
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var cell = xlApp.ActiveCell;
            string origFormula = cell.Formula;
        }

        /// <summary>
        /// Trying to create a button which generates over 4 lines:
        /// - Validated cell with material list
        /// - B 
        /// - H
        /// - Cross section Tag
        /// </summary>
        /// <param name="control"></param>
        public void CreatesCrossSection(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var activeCell = xlApp.ActiveCell;
            activeCell.Value2 = "Material = ";
            int row = activeCell.Row;
            int col = activeCell.Column;
            var ce = new ExcelReference(row,col);
            
            Range cell_1 = xlApp.Range[activeCell.Row + 1, activeCell.Column+1].Cells;
            cell_1.Value2 = "Test_Approved";
            //ValidateCellWithMaterials(xlApp.Range[activeCell.Row, activeCell.Column + 1]);
            //xlApp.Range[activeCell.Row + 1, activeCell.Column].Value2 = "width :";
            //xlApp.Range[activeCell.Row + 2, activeCell.Column].Value2 = "height :";
            //xlApp.Range[activeCell.Row + 3, activeCell.Column].Value2 = "CrossSection :";
            //xlApp.Range[activeCell.Row + 1, activeCell.Column + 1].Value2 = 100;
            //xlApp.Range[activeCell.Row + 2, activeCell.Column+2].Value2 = 200;
            //xlApp.Range[activeCell.Row + 3, activeCell.Column+3].Formula = String.Format("=CrossSectionTag({0},{1},{2})",
            //    xlApp.Range[activeCell.Row + 1, activeCell.Column + 1].Address,
            //    xlApp.Range[activeCell.Row + 2, activeCell.Column + 1].Address,
            //    xlApp.Range[activeCell.Row + 3, activeCell.Column + 1].Address);


            //ExcelReference b3 = new ExcelReference(2,1);
            //b3.SetValue("Hey! It works...");
            //


        }



        private void ValidateCellWithMaterials(Range cell)
        {
            var list = new System.Collections.Generic.List<string>();


            List<string> materials = new List<string>();
            foreach (var mtx in Enum.GetNames(typeof(MaterialTimberSoftwood.Grades)))
            {
                materials.Add(mtx);
            }


            var flatList = string.Join(",", materials.ToArray());
            string initialValue = materials[0];


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
