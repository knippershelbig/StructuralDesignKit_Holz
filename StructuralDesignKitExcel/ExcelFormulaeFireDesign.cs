using ExcelDna.Integration;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            [ExcelArgument(Description = "Length of the board fasteners")] int lengthFastener)

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



                def = EC5_Utilities.ComputeCharringDepthProtectedBeam(t, timber, Plasterboards, Thicknesses, closedJoint, lengthFastener);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }
            return def;
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


