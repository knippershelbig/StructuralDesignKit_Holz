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
using StructuralDesignKitLibrary.CrossSections.Interfaces;


namespace StructuralDesignKitExcel
{

    /// <summary>
    /// Methods linked to the ribbon buttons
    /// </summary>
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
                    writer.WriteAttributeString("screentip", method.CustomAttributes.ToList()[0].NamedArguments[1].TypedValue.Value.ToString());
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
            ExcelHelpers.WorkBookOpen(xlApp); //Ensure a workbook is open

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
            ExcelHelpers.WorkBookOpen(xlApp); //Ensure a workbook is open

            var baseCell = xlApp.ActiveCell;
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

            ValidateCellWithList(gradeCell, TimberTypes);
            ValidateCellWithList(ServiceClassCell, ServiceClasses);
            ValidateCellWithList(LoadDurationCell, LoadDurations);

            string formula = string.Format("=SDK.Factors.Kmod({0},{1},{2})", gradeCell.Address[false, false], ServiceClassCell.Address[false, false], LoadDurationCell.Address[false, false]);
            activeCell.Value2 = formula;

            //Format
            baseCell.Font.Bold = true;
            for (int i = 0; i < 5; i++)
            {
                baseCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            for (int i = 0; i < 3; i++)
            {
                baseCell.Offset[i + 2, 0].Interior.Color = XlRgbColor.rgbLightYellow;
            }

            var range = xlApp.Range[baseCell, baseCell.Offset[4, 0]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Columns.AutoFit();


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
            ExcelHelpers.WorkBookOpen(xlApp); //Ensure a workbook is open

            var baseCell = xlApp.ActiveCell;
            var activeCell = xlApp.ActiveCell;

            string formula = string.Empty;

            //----------------------------------
            //Define cross section and material
            //----------------------------------

            int blockCSLength = 10;
            //Captions
            activeCell.Value2 = "Cross section";
            activeCell = baseCell.Offset[1, 0]; activeCell.Value2 = "b";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "h";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Material";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Cross Section";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Service Class";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Buckling Length Ly";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Buckling Length Lz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Torsional buckling Length Ltb";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Member length in tension (LVL)";


            //Values
            activeCell = baseCell.Offset[1, 1]; activeCell.Value2 = "100";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "400";
            activeCell = activeCell.Offset[1, 0]; ValidateCellWithList(activeCell, ExcelHelpers.AllMaterialAsList());
            activeCell.Value2 = "GL24h";
            formula = string.Format("=SDK.Utilities.CreateRectangularCrossSection({0},{1},{2})", activeCell.Offset[-2, 0].Address[false, false], activeCell.Offset[-1, 0].Address[false, false], activeCell.Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = formula;
            activeCell = activeCell.Offset[1, 0]; ValidateCellWithList(activeCell, ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.ServiceClass>());
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "3";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "1.5";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "3";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "3";


            //Units - Comments
            activeCell = baseCell.Offset[1, 2]; activeCell.Value2 = "mm";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "mm";
            activeCell = activeCell.Offset[4, 0]; activeCell.Value2 = "m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "m";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "m";

            //Format
            baseCell.Font.Bold = true;
            activeCell = baseCell.Offset[1, 1];
            for (int i = 0; i < blockCSLength - 1; i++)
            {
                activeCell.Offset[i, 0].Interior.Color = XlRgbColor.rgbLightYellow;
                activeCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignRight;
            }
            var range = xlApp.Range[baseCell, baseCell.Offset[blockCSLength - 1, 2]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;



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
            baseCell.Offset[blockForceStart, 0].Font.Bold = true;
            activeCell = baseCell.Offset[blockForceStart + 1, 1];
            for (int i = 0; i < blockForceLength - 1; i++)
            {
                activeCell.Offset[i, 0].Interior.Color = XlRgbColor.rgbLightYellow;
                activeCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignRight;
            }
            range = xlApp.Range[baseCell.Offset[blockForceStart, 0], baseCell.Offset[blockForceStart + blockForceLength - 1, 2]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


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
            baseCell.Offset[blockStressesStart, 0].Font.Bold = true;
            activeCell = baseCell.Offset[blockStressesStart + 1, 1];
            for (int i = 0; i < blockStressesLength - 1; i++)
            {
                activeCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ((dynamic)activeCell.Offset[i, 0]).NumberFormatLocal = "0.00";
            }
            range = xlApp.Range[baseCell.Offset[blockStressesStart, 0], baseCell.Offset[blockStressesStart + blockStressesLength - 1, 2]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;





            //----------------------------------
            //Define Checks Factors
            //----------------------------------

            int blockFactorStart = 11;
            int BlockFactorLength = 11;

            //Captions
            activeCell = baseCell.Offset[blockFactorStart, 4]; activeCell.Value2 = "Factors";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kmod";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Ym";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Khy";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Khz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcy";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcz";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcrit";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kcr";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kh_tension";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kl_tension_LVL";

            //values
            formula = String.Format("=SDK.Factors.Kmod({0},{1},{2})",
                baseCell.Offset[3, 1].Address[false, false],
                baseCell.Offset[5, 1].Address[false, false],
                baseCell.Offset[blockForceStart + 1, 1].Address[false, false]);
            activeCell = baseCell.Offset[blockFactorStart + 1, 5]; activeCell.Formula = formula;
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Ym({0})", baseCell.Offset[3, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Bending({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[2, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Bending({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[1, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kc({0},{1}*1000,{2}*1000,0)", baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kc({0},{1}*1000,{2}*1000,1)", baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kcrit({0},{1}*1000)", baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[8, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kcr({0})", baseCell.Offset[3, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kh_Tension({0},{1})", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[2, 1].Address[false, false]);
            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.Factors.Kl_LVL({0},{1}*1000)", baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[8, 1].Address[false, false]);


            //Format
            baseCell.Offset[blockFactorStart, 4].Font.Bold = true;
            activeCell = baseCell.Offset[blockFactorStart + 1, 5];
            for (int i = 0; i < BlockFactorLength - 1; i++)
            {
                activeCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ((dynamic)activeCell.Offset[i, 0]).NumberFormatLocal = "0.00";
            }
            range = xlApp.Range[baseCell.Offset[blockFactorStart, 4], baseCell.Offset[blockFactorStart + BlockFactorLength - 1, 5]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            //----------------------------------
            //Define checks
            //----------------------------------
            int BlockCheckLength = 10;

            //Captions
            activeCell = baseCell.Offset[0, 4]; activeCell.Value2 = "Eurocodes 5 DIN EN 1995-1 Checks";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Tension Parallel To Grain_6.1.2: ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Compression Parallel To Grain_6.1.4: ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending_6.1.6 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Tension_6.2.3 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Compression_6.2.4 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Bending And Buckling_6.3.2 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Lateral Torsional buckling_6.3.3";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Shear_6.1.7 : ";
            activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Torsion_6.1.8 : ";

            //Value
            activeCell = baseCell.Offset[1, 5]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.TensionParallelToGrain_6.1.2({0},{1},{2},{3},{4},{5})",
                baseCell.Offset[blockStressesStart + 1, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[blockFactorStart + 1, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 2, 5].Address[false, false], baseCell.Offset[blockFactorStart + 8, 5].Address[false, false], baseCell.Offset[blockFactorStart + 9, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.CompressionParallelToGrain_6.1.4({0},{1},{2},{3})",
                baseCell.Offset[blockStressesStart + 2, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false], baseCell.Offset[blockFactorStart + 1, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 2, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Bending_6.1.6({0},{1},{2},{3},{4},{5},{6})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[4, 1].Address[false, false],
                baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 3, 5].Address[false, false], baseCell.Offset[blockFactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndTension_6.2.3({0},{1},{2},{3},{4},{5},{6},{7},{8},{9})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 1, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 3, 5].Address[false, false], baseCell.Offset[blockFactorStart + 4, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 8, 5].Address[false, false], baseCell.Offset[blockFactorStart + 9, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndCompression_6.2.4({0},{1},{2},{3},{4},{5},{6},{7})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 2, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 3, 5].Address[false, false], baseCell.Offset[blockFactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.BendingAndBuckling_6.3.2({0},{1},{2},{3}*1000,{4}*1000,{5},{6},{7},{8},{9})",
                baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 2, 1].Address[false, false],
                baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false], baseCell.Offset[4, 1].Address[false, false],
                baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 3, 5].Address[false, false], baseCell.Offset[blockFactorStart + 4, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.LateralTorsionalBuckling_6.3.3({0},{1},{2},{3}*1000,{4}*1000,{5}*1000,{6},{7},{8},{9},{10})",
                   baseCell.Offset[blockStressesStart + 6, 1].Address[false, false], baseCell.Offset[blockStressesStart + 7, 1].Address[false, false], baseCell.Offset[blockStressesStart + 2, 1].Address[false, false],
                baseCell.Offset[6, 1].Address[false, false], baseCell.Offset[7, 1].Address[false, false], baseCell.Offset[8, 1].Address[false, false], baseCell.Offset[4, 1].Address[false, false],
                baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false],
                baseCell.Offset[blockFactorStart + 3, 5].Address[false, false], baseCell.Offset[blockFactorStart + 4, 5].Address[false, false]);


            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Shear_6.1.7({0},{1},{2},{3},{4})",
                baseCell.Offset[blockStressesStart + 3, 1].Address[false, false], baseCell.Offset[blockStressesStart + 4, 1].Address[false, false], baseCell.Offset[3, 1].Address[false, false],
                baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false]);

            activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.CrossSectionChecks.Torsion_6.1.8({0},{1},{2},{3},{4},{5})",
                baseCell.Offset[blockStressesStart + 5, 1].Address[false, false], baseCell.Offset[blockStressesStart + 3, 1].Address[false, false], baseCell.Offset[blockStressesStart + 4, 1].Address[false, false],
                baseCell.Offset[4, 1].Address[false, false], baseCell.Offset[blockFactorStart + 1, 5].Address[false, false], baseCell.Offset[blockFactorStart + 2, 5].Address[false, false]);


            //Format
            baseCell.Offset[0, 4].Font.Bold = true;
            activeCell = baseCell.Offset[0, 5];
            for (int i = 0; i < BlockCheckLength; i++)
            {
                activeCell.Offset[i, 0].HorizontalAlignment = XlHAlign.xlHAlignRight;
                ((dynamic)activeCell.Offset[i, 0]).NumberFormatLocal = "0.0%";
                FormatCondition format = (FormatCondition)(activeCell.Offset[i, 0]).FormatConditions.Add(XlFormatConditionType.xlCellValue,
                                       XlFormatConditionOperator.xlGreater, 1);

                format.Font.Bold = true;
                format.Interior.Color = XlRgbColor.rgbRed;
            }
            range = xlApp.Range[baseCell.Offset[0, 4], baseCell.Offset[BlockCheckLength - 1, 5]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;






            //----------------------------------
            //Material properties
            //----------------------------------

            activeCell = baseCell.Offset[0, 7]; activeCell.Value2 = "Material Properties";


            MaterialProperties(control, baseCell.Offset[0, 7], baseCell.Offset[1, 1]);

            //------------------------------
            //Global Formating
            //------------------------------
            int j = blockStressesStart + blockStressesLength;
            var range1 = xlApp.Range[baseCell, baseCell.Offset[j, 0]];
            //Autofit column width
            for (int i = 0; i < 9; i++)
            {
                range1.Offset[0, i].Columns.AutoFit();
            }

            //Cell merging for block titles


            xlApp.Range[baseCell.Offset[blockForceStart, 0], baseCell.Offset[blockForceStart, 2]].Merge();
            baseCell.Offset[blockForceStart, 0].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            xlApp.Range[baseCell.Offset[blockStressesStart, 0], baseCell.Offset[blockStressesStart, 2]].Merge();
            baseCell.Offset[blockStressesStart, 0].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            xlApp.Range[baseCell.Offset[blockFactorStart, 4], baseCell.Offset[blockFactorStart, 5]].Merge();
            baseCell.Offset[blockFactorStart, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            xlApp.Range[baseCell.Offset[0, 4], baseCell.Offset[0, 5]].Merge();
            baseCell.Offset[0, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            xlApp.Range[baseCell, baseCell.Offset[0, 2]].Merge();
            baseCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        }

        /// <summary>
        /// Display all the material properties in a table 
        /// </summary>
        /// <param name="control">ribbon control</param>
        /// <param name="insertCell">cell where to insert the table. If null, book's active cell</param>
        /// <param name="csCell">first cell to look for the cross section. Format: b, h, Mat on top of each other</param>
        public void MaterialProperties(IRibbonControl control, Range insertCell = null, Range csCell = null)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            ExcelHelpers.WorkBookOpen(xlApp); //Ensure a workbook is open

            var baseCell = xlApp.ActiveCell;
            var activeCell = xlApp.ActiveCell;

            if (insertCell != null)
            {
                baseCell = insertCell;
                activeCell = insertCell;
            }

            var properties = typeof(StructuralDesignKitLibrary.Materials.IMaterialTimber).GetProperties().ToList();

            activeCell.Value2 = "Material Properties";
            activeCell.Offset[1, 0].Value2 = "b";
            activeCell.Offset[2, 0].Value2 = "h";
            activeCell.Offset[3, 0].Value2 = "Material";
            activeCell.Offset[4, 0].Value2 = "Cross Section";

            activeCell.Offset[1, 1].Value2 = 100;
            activeCell.Offset[2, 1].Value2 = 200;
            ValidateCellWithList(activeCell.Offset[3, 1], ExcelHelpers.AllMaterialAsList());
            activeCell.Offset[3, 1].Value2 = "GL24h";

            //Replace default values with linked values to a specific cross section:
            if (csCell != null)
            {
                activeCell.Offset[1, 1].Formula = string.Format("={0}", csCell.Address[false, false]);
                activeCell.Offset[2, 1].Formula = string.Format("={0}", csCell.Offset[1, 0].Address[false, false]);
                activeCell.Offset[3, 1].Formula = string.Format("={0}", csCell.Offset[2, 0].Address[false, false]);
            }

            activeCell.Offset[4, 1].Formula = string.Format("=SDK.Utilities.CreateRectangularCrossSection({0},{1},{2})",
                activeCell.Offset[1, 1].Address[false, false], activeCell.Offset[2, 1].Address[false, false], activeCell.Offset[3, 1].Address[false, false]);

            int count = 0;
            foreach (var prop in properties)
            {
                if (prop.Name != "Type")
                {
                    activeCell.Offset[5 + count, 0].Value2 = prop.Name;
                    activeCell.Offset[5 + count, 1].Formula = string.Format("=SDK.Material.Property({0}, \"{1}\")", activeCell.Offset[4, 1].Address[false, false], prop.Name);
                    count += 1;
                }

            }

            //Format
            baseCell.Font.Bold = true;
            for (int i = 0; i < properties.Count + 4; i++)
            {
                baseCell.Offset[i, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;
            }

            for (int i = 0; i < 3; i++)
            {
                baseCell.Offset[i + 1, 1].Interior.Color = XlRgbColor.rgbLightYellow;
            }

            var range = xlApp.Range[baseCell, baseCell.Offset[properties.Count + 3, 1]];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Columns.AutoFit();

            xlApp.Range[baseCell, baseCell.Offset[0, 1]].Merge();
            baseCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;


        }


        /// <summary>
        /// Create a cell validation with a given list
        /// </summary>
        /// <param name="cell">cell to modify</param>
        /// <param name="list">List to implement</param>
        private void ValidateCellWithList(Range cell, List<string> list)
        {
            if (cell == null) { throw new Exception("Please open a new workbook first"); }
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
