using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;
using Microsoft.Office.Interop.Excel;

namespace StructuralDesignKitExcel.RibbonActions
{
    internal class ConnectionButtonActions
    {

        public static void SteelConnectionButtonActions(Excel.Application xlApp, string name)
        {
            var baseCell = xlApp.ActiveCell;
            var activeCell = xlApp.ActiveCell;

            var LoadDurations = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.LoadDuration>();
            var TimberTypes = ExcelHelpers.AllMaterialAsList();
            var ServiceClasses = ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.EC5.EC5_Utilities.ServiceClass>();


            if (name == "SingleOuterSteelPlate")
            {
                //----INPUT----
                baseCell.Value2 = name;
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Timber 1";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Thickness t1";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Angle 1";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Plate thickness";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Diameter";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Fuk";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Rope Effect";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Fastener Tag";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Load Duration";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Service Class";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Kmod";



                Range Timber1 = baseCell.Offset[1, 1]; RibbonUtilities.ValidateCellWithList(Timber1, TimberTypes);
                Range Thick1 = Timber1.Offset[1, 0];
                Range Angle1 = Thick1.Offset[1, 0];
                Range Plate = Angle1.Offset[1, 0];
                Range Diameter = Plate.Offset[1, 0];
                Range Fuk = Diameter.Offset[1, 0];
                Range RopeEffect = Fuk.Offset[1, 0]; RibbonUtilities.ValidateCellWithList(RopeEffect, new List<string>() { "TRUE", "FALSE" });
                Range FastenerTag = RopeEffect.Offset[1, 0];
                Range LoadDuration = FastenerTag.Offset[1, 0]; RibbonUtilities.ValidateCellWithList(LoadDuration, LoadDurations);
                Range ServiceClass = LoadDuration.Offset[1, 0]; RibbonUtilities.ValidateCellWithList(ServiceClass, ServiceClasses);
                Range Kmod = ServiceClass.Offset[1, 0];


                Timber1.Value2 = "GL24h";
                Thick1.Value2 = 100;
                Angle1.Value2 = 0;
                Plate.Value2 = 6;
                Diameter.Value2 = 16;
                Fuk.Value2 = 400;
                RopeEffect.Value2 = true;
                FastenerTag.Formula = String.Format("=SDK.Utilities.CreateBoltTag({0},{1})", Diameter.Address[false, false], Fuk.Address[false, false]);
                LoadDuration.Value2 = "Permanent";
                ServiceClass.Value2 = "SC2";
                Kmod.Formula = string.Format("=SDK.Factors.Kmod({0},{1},{2})",
                    Timber1.Address[false, false], ServiceClass.Address[false, false], LoadDuration.Address[false, false]);


                //Results

                //Generate calculation to determine the failure mechanisms
                var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag.Value2);
                var timber1 = ExcelHelpers.GetTimberMaterialFromTag(Timber1.Value2);


                var shearCalculation = new SingleOuterSteelPlate(fastener, Plate.Value2, Angle1.Value2, timber1, Thick1.Value2, RopeEffect.Value2);



                //----Results
                activeCell = Kmod.Offset[1, -1]; activeCell.Value2 = "Fh1k";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Myrk";

                foreach (string failureMode in shearCalculation.FailureModes)
                {
                    activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Failure Mode " + failureMode;
                }

                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Characteristic Strength Rk";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "Design Capacity Rd";

                activeCell = Kmod.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.EC5.Connections.Fhk({0},{1},{2})",
                    FastenerTag.Address[false, false], Angle1.Address[false, false], Timber1.Address[false, false]);


                activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.EC5.Connections.Myrk({0})",
                    FastenerTag.Address[false, false]);

                foreach (string failureMode in shearCalculation.FailureModes)
                {
                    activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.EC5.Connection.SingleOuterSteelPlateFailureMode({0},{1},{2},{3},{4},{5},\"{6}\")",
                            FastenerTag.Address[false, false], Plate.Address[false, false], Angle1.Address[false, false], Timber1.Address[false, false], Thick1.Address[false, false],
                            RopeEffect.Address[false, false], failureMode);
                }

                activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("=SDK.EC5.Connection.SingleOuterSteelPlate({0},{1},{2},{3},{4},{5})",
                        FastenerTag.Address[false, false], Plate.Address[false, false], Angle1.Address[false, false], Timber1.Address[false, false], Thick1.Address[false, false],
                        RopeEffect.Address[false, false]);

                activeCell = activeCell.Offset[1, 0]; activeCell.Formula = string.Format("={0}*{1}/1.3/1000", activeCell.Offset[-1, 0].Address[false, false], Kmod.Address[false, false]);


                //Units
                activeCell = baseCell.Offset[2, 2]; activeCell.Value2 = "mm";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "°";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "mm";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "mm";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "N/mm²";
                activeCell = activeCell.Offset[6, 0]; activeCell.Value2 = "N/mm²";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "N.mm";

                foreach (string failureMode in shearCalculation.FailureModes)
                {
                    activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "N";
                }
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "N";
                activeCell = activeCell.Offset[1, 0]; activeCell.Value2 = "KN";


                //Formating
                baseCell.Font.Bold = true;
                activeCell = baseCell.Offset[1, 1];
                for (int i = 0; i < 10; i++)
                {
                    activeCell.Offset[i, 0].Interior.Color = XlRgbColor.rgbLightYellow;
                }


                activeCell = Kmod;
                for (int i = 0; i < 1; i++)
                {
                    activeCell = activeCell.Offset[1, 0];
                    ((dynamic)activeCell).NumberFormatLocal = ExcelHelpers.FormatSeparator(xlApp);
                }

                int count = shearCalculation.FailureModes.Count;

                for (int i = 0; i < count + 2; i++)
                {
                    activeCell = activeCell.Offset[1, 0];
                    ((dynamic)activeCell).NumberFormatLocal = "0";
                }

                   ((dynamic)activeCell.Offset[1, 0]).NumberFormatLocal = ExcelHelpers.FormatSeparator(xlApp);

                var range = xlApp.Range[baseCell, baseCell.Offset[15 + count, 2]];
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Columns.AutoFit();

                xlApp.Range[baseCell.Offset[0, 1], baseCell.Offset[15 + count, 1]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                baseCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xlApp.Range[baseCell, baseCell.Offset[0, 2]].Merge();
            }


        }




    }
}
