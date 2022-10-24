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
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;

namespace StructuralDesignKitExcel
{
    [ComVisible(true)] //To make Excel recognize the ribbon
    public class RibbonController : ExcelRibbon
    {

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




        public void OnButtonPressedKmod(IRibbonControl control)
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


        public void OnButtonPressedCSCheck(IRibbonControl control)
        {
            Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var activeCell = xlApp.ActiveCell;
            string formula = string.Empty;

            //Define cross section
            activeCell.Value2 = "Cross section";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "b";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "h";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Material";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Cross Section";

            activeCell = activeCell.Offset[-3, 1]; activeCell.Value2 = "100";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "400";
            activeCell = activeCell.Offset[1, 0]; ValidateCellWithList(activeCell, ExcelHelpers.AllMaterialAsList());
            activeCell.Value2 = "GL24h";
            formula = string.Format("=SDK.Material.CreateRectangularCrossSection({0},{1},{2})", activeCell.Offset[-2, 0].Address, activeCell.Offset[-1, 0].Address, activeCell.Address);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = formula;
            activeCell = activeCell.Offset[1, -1]; activeCell.Value2 = "Service Class";
            activeCell = activeCell.Offset[0, 1]; ValidateCellWithList(activeCell, ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.ServiceClass>());

            activeCell = activeCell.Offset[1, -1]; activeCell.Value2 = "Buckling Length Ly";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "3000";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "mm";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Buckling Length Lz";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "1500";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "mm";


            //Define forces
            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Loads";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Load duration";
            activeCell = activeCell.Offset[0, 1]; ValidateCellWithList(activeCell, ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.LoadDuration>());

            activeCell = activeCell.Offset[1, -1]; activeCell.Value2 = "Normal Force N = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "50";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Torsion Mx = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "2";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN.m";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Bending My = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "30";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN.m";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Bending Mz = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "5";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN.m";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Shear Vy = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "5";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN";

            activeCell = activeCell.Offset[1, -2]; activeCell.Value2 = "Shear Vz = ";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "15";
            activeCell = activeCell.Offset[0, 1]; activeCell.Value2 = "KN";

            activeCell = activeCell.Offset[2, -2]; activeCell.Value2 = "Kmod";
            formula = String.Format("=SDK.Factors.Kmod({0},{1},{2})",
                activeCell.Offset[-14, 1].Address,
                activeCell.Offset[-12, 1].Address,
                activeCell.Offset[-8, 1].Address);
            activeCell = activeCell.Offset[0, 1]; activeCell.Formula = formula;

            activeCell = activeCell.Offset[1, -1]; activeCell.Value2 = "Ym";
            activeCell = activeCell.Offset[0, 1]; activeCell.Formula = string.Format("=SDK.Factors.Ym({0})", activeCell.Offset[-15, 0].Address);

            var CrossSectionAdress = activeCell.Offset[-14, 0].Address;

            activeCell = activeCell.Offset[1, -1]; activeCell.Formula = "σ_N = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_Torsion = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_My = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_Mz = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_y = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_z = ";
            activeCell = activeCell.Offset[-5, 1]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.NormalForce({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.TorsionShear({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.BendingY({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.BendingY({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.ShearY({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.ShearZ({0},{1})",
                activeCell.Offset[-9, 0].Address, CrossSectionAdress);


            //SDK.CrossSection_StressCompute.BendingZ
            ////    SDK.CrossSection_StressCompute.NormalForce
            //    SDK.CrossSection_StressCompute.ShearY
            //    SDK.CrossSection_StressCompute.ShearZ
            //    SDK.CrossSection_StressCompute.TorsionShear


            //Define stresses


            //Define checks


            //Add colors - style

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
