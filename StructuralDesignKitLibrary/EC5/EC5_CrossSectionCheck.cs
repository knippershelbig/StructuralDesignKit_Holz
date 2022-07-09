using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StructuralDesignKitLibrary.EC5
{
    public static class EC5_CrossSectionCheck
    {




        /// <summary>
        /// Tension parallel to the grain EN 1995-1 §6.1.2 - Eq(6.1)
        /// </summary>
        /// <param name="Sig0_t_d">Design tensile stress</param>
        /// <param name="ft_0_k">Characteristic tensile strength</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="Kh">Size Factor for Cross section</param>
        /// <param name="Kl">Mofification factor for member Length</param>
        /// <returns>Design ratio for Tension parallel to the grain according to EN 1995-1 §6.1.2 - Eq(6.1)</returns>
        [Description("Tension parallel to the grain §6.1.2")]
        public static double TensionParallelToGrain(double Sig0_t_d, double ft_0_k, double Kmod, double Ym, double Kh = 1, double Kl = 1)
        {
            return Sig0_t_d / (Kh * Kl * ft_0_k * Kmod / Ym);
        }




        /// <summary>
        /// Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)
        /// </summary>
        /// <param name="Sig0_c_d">Design compressive stress</param>
        /// <param name="fc_0_k">Characteristic compressive strength</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for Compression parallel to the grain according to EN 1995-1 §6.1.4 - Eq(6.2)</returns>
        [Description("Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)")]
        public static double CompressionParallelToGrain(double Sig0_c_d, double fc_0_k, double Kmod, double Ym)
        {
            return Sig0_c_d / (fc_0_k * Kmod / Ym);
        }



        /// <summary>
        /// Bending EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12)
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="fmy_k">Characteristic bending strength in cross section Y axis</param>
        /// <param name="fmz_k">Characteristic bending strength in cross section Z axis</param>
        /// <param name="crossSection">Cross section Object to check</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <returns>Design ratio for Bending according to EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12) - Only the most onerous result of the two equations is returned</returns>
        [Description("Bending EN 1995-1 §6.1.4 - Eq(6.11) + Eq(6.12)")]
        public static double Bending(double SigMyd, double SigMzd, double fmy_k, double fmz_k, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1)
        {

            //km -> Factor considering re-distribution of bending stresses in a cross-section
            double km = 1;

            if (crossSection is CrossSectionRectangular)
            {
                IMaterialTimber timber;
                if (material is IMaterialTimber)
                {
                    timber = (IMaterialTimber)material;

                    if (timber.Type == EC5_Utilities.TimberType.Softwood ||
                        timber.Type == EC5_Utilities.TimberType.Hardwood ||
                        timber.Type == EC5_Utilities.TimberType.Glulam ||
                        timber.Type == EC5_Utilities.TimberType.LVL ||
                        timber.Type == EC5_Utilities.TimberType.Baubuche)
                    {
                        km = 0.7;
                    }
                }
            }

            double eq_6_11 = SigMyd / (fmy_k * khy * Kmod / Ym) + km * SigMzd / (fmz_k * khz * Kmod / Ym);
            double eq_6_12 = km * SigMyd / (fmy_k * khy * Kmod / Ym) + SigMzd / (fmz_k * khz * Kmod / Ym);

            return Math.Max(eq_6_11, eq_6_12);
        }


        /// <summary>
        /// Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)
        /// </summary>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="fv_k">Characteristic shear strength</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for shear according to DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54) - Only the most onerous result is returned</returns>
        [Description("Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)")]
        public static double Shear(double TauYd, double TauZd, double fv_k, IMaterial material, double Kmod, double Ym)
        {
            List<double> ShearResults = ComputeShearCheck(TauYd, TauZd, fv_k, material, Kmod, Ym);

            var sortedShearResults = ShearResults.OrderByDescending(p => p).ToList();

            return sortedShearResults[0];
        }

        /// <summary>
        /// Computes the shear checks but returns a list of doubles for the 3 equations. Can be used both for shear check and torsion check
        /// </summary>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="fv_k">Characteristic shear strength</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Returns a List with 3 values : ratioShearY, ratioShearZ and  CombinedShear Y² + Z² </returns>
        private static List<double> ComputeShearCheck(double TauYd, double TauZd, double fv_k, IMaterial material, double Kmod, double Ym)
        {
            double kcr = EC5_Factors.Kcr(material);

            double ratioShearY = TauYd / (kcr * fv_k * Kmod / Ym);

            double ratioShearZ = TauZd / (kcr * fv_k * Kmod / Ym);

            //Additional check from DIN EN 1995-1 NA-DE to 6.1.7 -> Eq NA.54
            double CombinedShear = Math.Pow(ratioShearY, 2) + Math.Pow(ratioShearZ, 2);

            return new List<double>() { ratioShearY, ratioShearZ, CombinedShear };
        }




        //Torsion
        /// <summary>
        /// Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)
        /// </summary>
        /// <param name="TauTorsion">Design torsion shear stress</param>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="fv_k">Characteristic shear strength</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for Torsion and shear according to DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55) </returns>
        /// <exception cref="Exception"></exception>
        [Description("Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)")]
        public static double Torsion(double TauTorsion, double TauYd, double TauZd, double fv_k, ICrossSection crossSection, IMaterial material, double Kmod, double Ym)
        {
            var combinedShear = ComputeShearCheck(TauYd, TauZd, fv_k, material, Kmod, Ym)[2];

            double Kshape = 0;

            if (crossSection is CrossSectionRectangular)
            {
                CrossSectionRectangular rectCS = (CrossSectionRectangular)crossSection;
                Kshape = Math.Min(1 + 0.15 * rectCS.H / rectCS.B, 2);
            }
            else throw new Exception("Currently only Rectangular Cross sections are covered for torsion check");


            return TauTorsion / (Kshape * (fv_k * Kmod / Ym)) + combinedShear;
        }





        //Bending and tension
        //Bending and compression
        //Instabilities







    }
}
