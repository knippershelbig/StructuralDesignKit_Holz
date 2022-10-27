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
using StructuralDesignKitLibrary.EC5;

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

            string formula = string.Format("=SDK.Factors.Kmod({0},{1},{2})", gradeCell.Address[false, false], ServiceClassCell.Address[false, false], LoadDurationCell.Address[false, false]);
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
            var baseCell = xlApp.ActiveCell;
            var activeCell = xlApp.ActiveCell;

            string formula = string.Empty;

            //----------------------------------
            //Define cross section and material
            //----------------------------------

            int blockCSLength = 9;
            //Captions
            activeCell.Value2 = "Cross section";
            activeCell = baseCell.Offset[1, 0]; activeCell.Value2 = "b";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "h";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Material";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Cross Section";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Service Class";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Buckling Length Ly";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Buckling Length Lz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Member length in tension (LVL)";


            //Values
            activeCell = baseCell.Offset[1, 1]; activeCell.Value2 = "100";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "400";
            activeCell = activeCell.Offset[1, 0]; ValidateCellWithList(activeCell, ExcelHelpers.AllMaterialAsList());
            activeCell.Value2 = "GL24h";
            formula = string.Format("=SDK.Material.CreateRectangularCrossSection({0},{1},{2})", activeCell.Offset[-2, 0].Address[false, false], activeCell.Offset[-1, 0].Address[false, false], activeCell.Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = formula;
            activeCell = activeCell.Offset[1, 0]; ValidateCellWithList(activeCell, ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.ServiceClass>());
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "3";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "1.5";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "3";


            //Units - Comments
            activeCell = baseCell.Offset[1, 2]; activeCell.Value2 = "mm";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "mm";
            activeCell = activeCell.Offset[4, 0]; activeCell.Value2 = "m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "m";

            //Format

            //----------------------------------
            //Define forces/loads
            //----------------------------------
            int blockForceStart = blockCSLength + 1;
            int blockForceLength = 9;
            //Captions
            activeCell = baseCell.Offset[blockForceStart, 0]; activeCell.Value2 = "Loads";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Load duration";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Normal tension Force N(+) = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Normal Compression Force N(-) = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Shear Vy = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Shear Vz = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Torsion Mx = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending My = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending Mz = ";

            //Values
            activeCell = baseCell.Offset[blockForceStart + 1, 1]; ValidateCellWithList(activeCell, ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.LoadDuration>());
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "50";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "-100";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "5";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "5";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "2";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "30";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "15";

            //Units - Comments
            activeCell = baseCell.Offset[blockForceStart + 2, 2]; activeCell.Value2 = "KN";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN.m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN.m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN.m";



            //Format


            //----------------------------------
            //Define Stresses
            //----------------------------------
            int blockStressesStart = blockForceStart + blockForceLength + 1;
            int blockStressesLength = 8;

            //Captions
            activeCell = baseCell.Offset[blockStressesStart, 0]; activeCell.Formula = "Stresses";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_N_Tension = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_N_Comp = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_y = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_z = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "τ_Torsion = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_My = ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = "σ_Mz = ";


            //Values
            var CrossSectionAdress = baseCell.Offset[4, 1].Address[false, false];

            activeCell = baseCell.Offset[blockStressesStart + 1, 1]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.NormalForce({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.NormalForce({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.ShearY({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.ShearZ({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.TorsionShear({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.BendingY({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = String.Format("=SDK.CrossSection_StressCompute.BendingZ({0},{1})",
                activeCell.Offset[-9, 0].Address[false, false], CrossSectionAdress);


            //Units-Comments
            for (int i = 0; i < 7; i++)
            {
                baseCell.Offset[blockStressesStart + 1 + i, 2].Value2 = "N/mm²";
            }
            //Format


            //----------------------------------
            //Define Checks Factors
            //----------------------------------

            int FactorStart = 10;

            //Captions
            activeCell = baseCell.Offset[FactorStart, 4]; activeCell.Value2 = "Factors";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kmod";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Ym";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Khy";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Khz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcy";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcr";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kh_tension";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kl_tension_LVL";

            //values
            formula = String.Format("=SDK.Factors.Kmod({0},{1},{2})",
                baseCell.Offset[3, 1].Address[false, false],
                baseCell.Offset[5, 1].Address[false, false],
                baseCell.Offset[blockForceStart + 1, 1].Address[false, false]);
            activeCell = baseCell.Offset[FactorStart + 1, 5]; activeCell.Formula = formula;
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Ym({0})", baseCell.Offset[3, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Bending({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[2, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Bending({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[1, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kc({0},{1},{2}*1000,0)", baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kc({0},{1},{2}*1000,1)", baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kcr({0})", baseCell.Offset[3, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Tension({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[2, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kl_LVL({0},{1}*1000)", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[8, 1].Address[false, false]);


            //Format


            //----------------------------------
            //Define checks
            //----------------------------------

            //Captions
            activeCell = baseCell.Offset[0, 4]; activeCell.Value2 = "Eurocodes 5 Checks";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Tension Parallel To Grain_6.1.2: ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Compression Parallel To Grain_6.1.4: ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending_6.1.6 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Tension_6.2.3 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Compression_6.2.4 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Buckling_6.3.2 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Shear_6.1.7 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Torsion_6.1.8 : ";

            //Value
            activeCell = baseCell.Offset[1, 5]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.TensionParallelToGrain_6.1.2({0},{1},{2},{3},{4},{5})",
                baseCell.Offset[blockStressesStart + 1, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[FactorStart + 1, 5].Address[false, false],
                baseCell.Offset[FactorStart + 2, 5].Address[false, false], baseCell.Offset[FactorStart + 8, 5].Address[false, false], baseCell.Offset[FactorStart + 9, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.CompressionParallelToGrain_6.1.4({0},{1},{2},{3})",
                baseCell.Offset[blockStressesStart + 2, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[FactorStart + 1, 5].Address[false, false],
                baseCell.Offset[FactorStart + 2, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Bending_6.1.6({0},{1},{2},{3},{4},{5},{6})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[4, 1].Address[false, false],
                baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false],
                baseCell.Offset[FactorStart + 3, 5].Address[false, false], baseCell.Offset[FactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndTension_6.2.3({0},{1},{2},{3},{4},{5},{6},{7},{8},{9})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 1, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false],
                baseCell.Offset[FactorStart + 3, 5].Address[false, false], baseCell.Offset[FactorStart + 4, 5].Address[false, false],
                baseCell.Offset[FactorStart + 8, 5].Address[false, false], baseCell.Offset[FactorStart + 9, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndCompression_6.2.4({0},{1},{2},{3},{4},{5},{6},{7})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 2, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false],
                baseCell.Offset[FactorStart + 3, 5].Address[false, false], baseCell.Offset[FactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndBuckling_6.3.2({0},{1},{2},{3}*1000,{4}*1000,{5},{6},{7},{8},{9})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 2, 1].Address[false, false],
                baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false], baseCell.Offset[4, 1].Address[false, false],
                baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false],
                baseCell.Offset[FactorStart + 3, 5].Address[false, false], baseCell.Offset[FactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Shear_6.1.7({0},{1},{2},{3},{4})",
                baseCell.Offset[blockStressesStart + 3, 1].Address[false, false], baseCell.Offset[blockStressesStart + 4, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false],
                baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Torsion_6.1.8({0},{1},{2},{3},{4},{5})",
                baseCell.Offset[blockStressesStart + 5, 1].Address[false, false], baseCell.Offset[blockStressesStart + 3, 1].Address[false, false], baseCell.Offset[blockStressesStart + 4, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[FactorStart + 1, 5].Address[false, false], baseCell.Offset[FactorStart + 2, 5].Address[false, false]);


            //Format

            //----------------------------------
            //Material properties
            //----------------------------------

            activeCell = baseCell.Offset[0, 7]; activeCell.Value2 = "Material Properties";

            //Caption
            //Define list of material properties
            var properties = typeof(StructuralDesignKitLibrary.Materials.IMaterialTimber).GetProperties().ToList();
            int count = 0;

            foreach (var prop in properties)
            {
                if (prop.Name != "Type")
                {
                    baseCell.Offset[1 + count, 7].Value2 = prop.Name;
                    baseCell.Offset[1 + count, 8].Formula = string.Format("=SDK.Material.Property({0}, \"{1}\")", baseCell.Offset[4, 1].Address[false, false], prop.Name);
                    count +=1;
                }

            }
            for (int i = 0; i < properties.Count; i++)
            {
                
            }
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
            //    xlApp.Range[activeCell.Row + 1, activeCell.Column + 1].Address[false,false],
            //    xlApp.Range[activeCell.Row + 2, activeCell.Column + 1].Address[false,false],
            //    xlApp.Range[activeCell.Row + 3, activeCell.Column + 1].Address[false,false]);


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
