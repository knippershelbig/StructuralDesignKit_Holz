using ExcelDna.Integration;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;

namespace StructuralDesignKitExcel
{
    public static partial class ExcelFormulae
    {
        #region Fire Design
        //-------------------------------------------
        //Charring depth beam unprotected
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the charring depth for an unprotected beam in [mm]",
            Name = "SDK.FireDesign.ComputeCharringDepthUnprotectedBeam",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeCharringDepthUnprotectedBeam(
            [ExcelArgument(Description = "Fire duration in minutes")] int t,
            [ExcelArgument(Description = "Material Tag")] string material)
        {
            double def = -1;

            try
            {
                var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
                def = EC5_Utilities.ComputeCharringDepthUnprotectedBeam(t, timber);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return def;
        }




        //-------------------------------------------
        //Charring depth Protected beam 
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the charring depth for a protected beam in [mm]",
            Name = "SDK.FireDesign.ComputeCharringDepthProtectedBeam",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeCharringDepthProtectedBeam(
            [ExcelArgument(Description = "Fire duration in minutes")] int t,
            [ExcelArgument(Description = "Material Tag")] string material,
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2,
            [ExcelArgument(Description = "Joints between boards < 2mm")] bool closedJoint,
            [ExcelArgument(Description = "if the protection considered is horizontal -> true, otherwise -> false (vertical)")] bool horizontal)

        {
            double def = -1;

            try
            {
                var timber = ExcelHelpers.GetTimberMaterialFromTag(material);

                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }



                def = EC5_Utilities.ComputeCharringDepthProtectedBeam(t, timber, Plasterboards, Thicknesses, closedJoint, horizontal);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return def;
        }




        [ExcelFunction(Description = "Compute start of combustion (in minutes) of a protected surface according to EN 1995-1-2 §3.4.3.3",
            Name = "SDK.FireDesign.ComputeCombustionStart",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeCombustionStart(
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2,
            [ExcelArgument(Description = "Joints between boards < 2mm")] bool closedJoint)

        {
            double tch = -1;

            try
            {
                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }

                tch = EC5_Utilities.ComputeCombustionStart(Plasterboards, Thicknesses, closedJoint);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return tch;
        }




        [ExcelFunction(Description = "Compute the Panel Failure Time (in minutes) according to the Austrian approach",
            Name = "SDK.FireDesign.ComputePanelFailureTime",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputePanelFailureTime(
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2,
            [ExcelArgument(Description = "Joints between boards < 2mm")] bool closedJoint,
            [ExcelArgument(Description = "if the protection considered is horizontal -> true, otherwise -> false (vertical)")] bool horizontal)

        {
            double tf = -1;
            try
            {
                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }

                double tch = EC5_Utilities.ComputeCombustionStart(Plasterboards, Thicknesses, closedJoint);
                tf = EC5_Utilities.ComputePanelFailureTime(tch, Plasterboards, Thicknesses, horizontal);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return tf;
        }



        [ExcelFunction(Description = "Insulation coefficient according to DIN EN 1995-1-2 §3.4.3.2 (2)",
            Name = "SDK.FireDesign.ComputeK2",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeK2(
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2)
        {
            double k2 = -1;
            try
            {

                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }
                k2 = EC5_Utilities.ComputeK2(Plasterboards, Thicknesses);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return k2;
        }



        [ExcelFunction(Description = "Compute the time limit (in minutes) above which the charring rate gets back to Bn",
            Name = "SDK.FireDesign.ComputeTa",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeTa(
            [ExcelArgument(Description = "Material Tag")] string material,
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2,
            [ExcelArgument(Description = "Joints between boards < 2mm")] bool closedJoint,
            [ExcelArgument(Description = "if the protection considered is horizontal -> true, otherwise -> false (vertical)")] bool horizontal)

        {
            double ta = -1;
            try
            {
                var timber = ExcelHelpers.GetTimberMaterialFromTag(material);

                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }


                double tch = EC5_Utilities.ComputeCombustionStart(Plasterboards, Thicknesses, closedJoint);
                double tf = EC5_Utilities.ComputePanelFailureTime(tch, Plasterboards, Thicknesses, horizontal);
                double k2 = EC5_Utilities.ComputeK2(Plasterboards, Thicknesses);
                ta = EC5_Utilities.ComputeTa(tch, tf, k2, timber.Bn);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return ta;
        }



        [ExcelFunction(Description = "Return the minimum fire protection fastener length in [mm] according to EN 1995-1-2 §3.4.3.4 (4) - Eq 3.16",
            Name = "SDK.FireDesign.ComputeProtectionMinFastenerLength",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeProtectionMinFastenerLength(
            [ExcelArgument(Description = "fastener diameter in [mm]")] double d,
            [ExcelArgument(Description = "Fire duration in minutes")] int t,
            [ExcelArgument(Description = "Material Tag")] string material,
            [ExcelArgument(Description = "Protection board 1 type (external if 2 boards)")] string board1,
            [ExcelArgument(Description = "Protection board 1 thickness [mm] (external if 2 boards)")] double boardThickness1,
            [ExcelArgument(Description = "Protection board 2 type  (internal if 2 boards)")] string board2,
            [ExcelArgument(Description = "Protection board 2 thickness [mm] (internal if 2 boards)")] double boardThickness2,
            [ExcelArgument(Description = "Joints between boards < 2mm")] bool closedJoint,
            [ExcelArgument(Description = "if the protection considered is horizontal -> true, otherwise -> false (vertical)")] bool horizontal)
        {
            double lmin = -1;

            try
            {
                var timber = ExcelHelpers.GetTimberMaterialFromTag(material);

                List<PlasterboardType> Plasterboards = new List<PlasterboardType>();
                List<double> Thicknesses = new List<double>();
                if (board1 != "none" && boardThickness1 != 0)
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness1);
                }

                if (boardThickness2 != 0 && board2 != "none")
                {
                    Plasterboards.Add(ExcelHelpers.GetPlaterBoardTypeFromString(board1));
                    Thicknesses.Add(boardThickness2);
                }

                lmin = EC5_Utilities.ComputeProtectionMinFastenerLength(d, t, timber, Plasterboards, Thicknesses, closedJoint, horizontal);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return lmin;
        }

    

        //-------------------------------------------
        //Reduced cross section
        //-------------------------------------------
        [ExcelFunction(Description = "Create a cross section Tag with reduced cross section dur to fire charring",
            Name = "SDK.FireDesign.CreateReducedCrossSection",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static string CreateReducedCrossSection(string initialCrossSection, int t, bool top, bool bottom, bool left, bool right)
        {
            var initCS = ExcelHelpers.CreateRectangularCrossSection(initialCrossSection);
            var reducedCS = initCS.ComputeReducedCrossSection(t, top, bottom, left, right);

            return ExcelHelpers.CreateRectangularCrossSectionTag(reducedCS.B, reducedCS.H, (IMaterialTimber)reducedCS.Material);
        }

        #endregion
    }
}


