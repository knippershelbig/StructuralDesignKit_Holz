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
using System.Xml;
using System.IO;

namespace StructuralDesignKitExcel
{
    [ComVisible(true)] //To make Excel recognize the ribbon
    public class RibbonController : ExcelRibbon
    {

        //    public override string GetCustomUI(string RibbonID)
        //    {
        //        return @" 
        //<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
        //  <ribbon>
        //    <tabs>
        //      <tab id='SDK' label='SDK Timber'>
        //        <group id='group1' label='My Group'>
        //          <button id='button1' label='My Button' onAction='OnButtonPressed'/>
        //          <button id='button2' label='2nd Fumction' onAction='OnFormulaFunction' tag='SDK.Materials.Fmk'/>
        //          <button id='CrossSectionTag' label='Creates a SDK Cross section Tag' onAction='CreatesCrossSection'/>
        //        </group >
        //      </tab>
        //    </tabs>
        //  </ribbon>
        //</customUI>";
        //    }


        public string GetMenuContent(IRibbonControl control)
        {

            string category = control.Tag;

            var menuXML = new StringWriter();
            using (XmlWriter writer = XmlWriter.Create(menuXML))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("menu", @"http://schemas.microsoft.com/office/2006/01/customui");
                writer.WriteAttributeString("xmlns", @"http://schemas.microsoft.com/office/2006/01/customui");

                //Get methods from a class
                var methods = typeof(ExcelFormulae).GetMethods();

                //filtering Factors Methods
                List<System.Reflection.MethodInfo> methodsWithArgument = new List<System.Reflection.MethodInfo>();
                foreach (var method in methods)
                {
                    if (method.CustomAttributes.ToList().Count >= 1)
                    {
                        if (method.CustomAttributes.ToList()[0].NamedArguments.Count >= 3)
                        {
                            methodsWithArgument.Add(method);
                        }
                    }
                }

                var categoryMethods = methodsWithArgument.Where(p => p.CustomAttributes.ToList()[0].NamedArguments[2].TypedValue.Value.ToString() == category).ToList();

                foreach (var method in categoryMethods)
                {
                    writer.WriteStartElement("button");
                    writer.WriteAttributeString("id", "button_" + method.Name);
                    writer.WriteAttributeString("label", method.Name);
                    writer.WriteAttributeString("onAction", "OnButton");
                    writer.WriteAttributeString("tag", method.CustomAttributes.ToList()[0].NamedArguments[0].TypedValue.Value.ToString());
                    writer.WriteEndElement();
                }

                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();

            }
            return menuXML.ToString();
        }




        public void OnButton(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var cell = xlApp.ActiveCell;

            // AppActivate Application.Caption;
            cell.Select();
            xlApp.SendKeys("{F2}");
            cell.Value2 = "=" + control.Tag;
            xlApp.SendKeys("{(}");

        }




        public void KmodTable(IRibbonControl control)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var activeCell = xlApp.ActiveCell;

            var LoadDurations = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.LoadDuration>();
            var TimberTypes = ExcelHelpers.AllMaterialAsList();
            var ServiceClasses = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.ServiceClass>();

            activeCell.Value2 = "Kmod";
            activeCell = activeCell.Offset[1, 0];
            //Timber grade service class laod duration
            var gradeCell = activeCell.Offset[1, 0];
            var ServiceClassCell = gradeCell.Offset[1, 0];
            var LoadDurationCell = ServiceClassCell.Offset[1, 0];

            //activeCell = activeCell.Offset[1, 0];
            ValidateCellWithList(gradeCell, TimberTypes);
            //activeCell = activeCell.Offset[1, 0];
            ValidateCellWithList(ServiceClassCell, ServiceClasses);
            //activeCell = activeCell.Offset[1, 0];
            ValidateCellWithList(LoadDurationCell, LoadDurations);
            
            string formula = string.Format("=SDK.Factors.Kmod({0},{1},{2})", gradeCell.Address, ServiceClassCell.Address, LoadDurationCell.Address);
            activeCell.Value2 = formula;
        }



        public void MaterialList(IRibbonControl control)
        {
            List<string> timberType = new List<string>();

            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var activeCell = xlApp.ActiveCell;
            switch (control.Tag)
            {
                case "Softwood":
                    timberType = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberSoftwood.Grades>();
                    break;
                case "Hardwood":
                    timberType = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberHardwood.Grades>();
                    break;
                case "Glulam":
                    timberType = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberGlulam.Grades>();
                    break;
                case "Baubuche":
                    timberType = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberBaubuche.Grades>();
                    break;
                case "All":
                    timberType = ExcelHelpers.AllMaterialAsList();
                    break;

            }

            ValidateCellWithList(activeCell, timberType);
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
            var ce = new ExcelReference(row, col);

            Range cell_1 = xlApp.Range[activeCell.Row + 1, activeCell.Column + 1].Cells;
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


        /// <summary>
        /// Create a cell validation with a given list
        /// </summary>
        /// <param name="cell">cell to modify</param>
        /// <param name="list">List to implement</param>
        private void ValidateCellWithList(Range cell, List<string> list)
        {

            var flatList = string.Join(",", list.ToArray());
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
