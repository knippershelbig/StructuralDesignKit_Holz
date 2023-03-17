using ExcelDna.Integration;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Windows;

namespace StructuralDesignKitExcel
{
    public static partial class ExcelFormulae
    {
        #region Fire Design
        //-------------------------------------------
        //Charring depth
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the charring depth for a beam in [mm]",
            Name = "SDK.FireDesign.ComputeCharringDepthUnprotected",
            IsHidden = false,
            Category = "SDK.FireDesign")]
        public static double ComputeCharringDepthUnprotected(
            [ExcelArgument(Description = "Fire duration in minutes")] int t,
            [ExcelArgument(Description = "Material Tag")] string material)
        {
            double def = -1;

            try
            {
                var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
                def = EC5_Utilities.ComputeCharringDepthUnprotected(t, timber);
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
        [ExcelFunction(Description="Create a cross section Tag with reduced cross section dur to fire charring",
            Name = "SDK.FireDesign.CreateReducedCrossSection",
            IsHidden =false,
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


