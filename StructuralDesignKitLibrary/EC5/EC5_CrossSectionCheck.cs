using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using StructuralDesignKitLibrary.CrossSections;
using System.Security.Policy;

namespace StructuralDesignKitLibrary.EC5
{
    public static class EC5_CrossSectionCheck
    {

        /// <summary>
        /// Tension parallel to the grain EN 1995-1 §6.1.2 - Eq(6.1)
        /// </summary>
        /// <param name="Sig0_t_d">Design tensile stress</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="Kh">Size Factor for Cross section</param>
        /// <param name="Kl_LVL">Mofification factor for LVL member length</param>
        /// <returns>Design ratio for Tension parallel to the grain according to EN 1995-1 §6.1.2 - Eq(6.1)</returns>
        [Description("Tension parallel to the grain §6.1.2")]
        public static double TensionParallelToGrain(double Sig0_t_d, IMaterial material, double Kmod, double Ym, double Kh = 1, double Kl_LVL = 1, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }


            var timber = (IMaterialTimber)material;
            double ft0_k = timber.Ft0k;

            if (timber.Type != EC5_Utilities.TimberType.LVL && timber.Type != EC5_Utilities.TimberType.Baubuche) Kl_LVL = 1;

            return Sig0_t_d / (Kh * Kl_LVL * ft0_k * Kmod / Ym * kfi);
        }


        /// <summary>
        /// Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)
        /// </summary>
        /// <param name="Sig0_c_d">Design compressive stress</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for Compression parallel to the grain according to EN 1995-1 §6.1.4 - Eq(6.2)</returns>
        [Description("Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)")]
        public static double CompressionParallelToGrain(double Sig0_c_d, IMaterial material, double Kmod, double Ym, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");
            double kfi = 1;

            //Fire factors
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fc0_k = timber.Fc0k;



            return Sig0_c_d / (fc0_k * Kmod / Ym * kfi);
        }


        /// <summary>
        /// Compression stresses at an angle to the grain 
        /// </summary>
        /// <param name="SigAlpha_c_d">Design compressive stress</param>
        /// <param name="angleToGrain">stress angle to the grain in Degree</param>
        /// <param name="material"></param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="kc90">factor taking into account the effect of stresses perpendicular to the grain</param>
        /// <returns></returns>
        [Description("Compression stresses at an angle to the grain EN 1995-1 §6.2.2 - Eq(6.16)")]
        public static double CompressionAtAnAngleToGrain(double SigAlpha_c_d, double angleToGrain, IMaterial material, double Kmod, double Ym, double kc90 = 1, bool FireCheck = false)
        {
            var timber = CheckMaterialTimber(material);

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            double angleRad = angleToGrain * Math.PI / 180;
            double fc0_d = timber.Fc0k * Kmod / Ym * kfi;
            double fc90_d = timber.Fc90k * Kmod / Ym * kfi;

            return SigAlpha_c_d / (fc0_d / ((fc0_d / (kc90 * fc90_d)) * Math.Pow(Math.Sin(angleRad), 2) + Math.Pow(Math.Cos(angleRad), 2)));
        }


        /// <summary>
        /// Bending EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12)
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="crossSection">Cross section Object to check</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <returns>Design ratio for Bending according to EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12) - Only the most onerous result of the two equations is returned</returns>
        [Description("Bending EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12)")]
        public static double Bending(double SigMyd, double SigMzd, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fmy_k = timber.Fmyk;
            double fmz_k = timber.Fmzk;


            double km = EC5_Factors.Km(crossSection, material);

            double eq_6_11 = SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + km * SigMzd / (fmz_k * khz * Kmod / Ym * kfi);
            double eq_6_12 = km * SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + SigMzd / (fmz_k * khz * Kmod / Ym * kfi);

            return Math.Max(eq_6_11, eq_6_12);
        }


        /// <summary>
        /// Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)
        /// </summary>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for shear according to DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54) - Only the most onerous result is returned</returns>
        [Description("Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)")]
        public static double Shear(double TauYd, double TauZd, IMaterial material, double Kmod, double Ym, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;

            List<double> ShearResults = ComputeShearCheck(TauYd, TauZd, material, Kmod, Ym);

            var sortedShearResults = ShearResults.OrderByDescending(p => p).ToList();

            return sortedShearResults[0];
        }


        /// <summary>
        /// Computes the shear checks but returns a list of doubles for the 3 equations. Can be used both for shear check and torsion check
        /// </summary>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="material">Material Object to check</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Returns a List with 3 values : ratioShearY, ratioShearZ and  CombinedShear Y² + Z² </returns>
        [Description("Computes the shear checks but returns a list of doubles for the 3 equations. Can be used both for shear check and torsion check")]
        private static List<double> ComputeShearCheck(double TauYd, double TauZd, IMaterial material, double Kmod, double Ym, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fv_k = timber.Fvk;


            double kcr = EC5_Factors.Kcr(material);

            double ratioShearY = TauYd / (kcr * fv_k * Kmod / Ym * kfi);

            double ratioShearZ = TauZd / (kcr * fv_k * Kmod / Ym * kfi);

            //Additional check from DIN EN 1995-1 NA-DE to 6.1.7 -> Eq NA.54
            double CombinedShear = Math.Pow(ratioShearY, 2) + Math.Pow(ratioShearZ, 2);

            return new List<double>() { ratioShearY, ratioShearZ, CombinedShear };
        }


        /// <summary>
        /// Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)
        /// </summary>
        /// <param name="TauTorsion">Design torsion shear stress</param>
        /// <param name="TauYd">Design shear stress on Y</param>
        /// <param name="TauZd">Design shear stress on Z</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <returns>Design ratio for Torsion and shear according to DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55) </returns>
        /// <exception cref="Exception"></exception>
        [Description("Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)")]
        public static double Torsion(double TauTorsion, double TauYd, double TauZd, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fv_k = timber.Fvk;


            var combinedShear = ComputeShearCheck(TauYd, TauZd, material, Kmod, Ym)[2];

            double Kshape = EC5_Factors.KShape(crossSection, material);


            return TauTorsion / (Kshape * (fv_k * Kmod / Ym * kfi)) + combinedShear;
        }


        /// <summary>
        /// Bending and tension EN 1995-1 §6.2.3 - Eq(6.17) + Eq(6.18)
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="Sig0_t_d">Design tensile stress</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <param name="Kh_Tension">Size Factor for Cross section in tension</param>
        /// <param name="Kl_LVL">Mofification factor for member Length</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        [Description("Bending and tension EN 1995-1 +NA §6.2.3 Eq(6.17) + Eq(6.18)")]
        public static double BendingAndTension(double SigMyd, double SigMzd, double Sig0_t_d, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1, double Kh_Tension = 1, double Kl_LVL = 1, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fmy_k = timber.Fmyk;
            double fmz_k = timber.Fmzk;
            double ft0_k = timber.Ft0k;


            double km = EC5_Factors.Km(crossSection, material);

            if (timber.Type != EC5_Utilities.TimberType.LVL && timber.Type != EC5_Utilities.TimberType.Baubuche) Kl_LVL = 1;

            double tensionRatio = Sig0_t_d / (Kh_Tension * Kl_LVL * ft0_k * Kmod / Ym * kfi);
            double eq_6_17 = tensionRatio + SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + km * SigMzd / (fmz_k * khz * Kmod / Ym * kfi);
            double eq_6_18 = tensionRatio + km * SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + SigMzd / (fmz_k * khz * Kmod / Ym * kfi);

            return Math.Max(eq_6_17, eq_6_18);
        }


        /// <summary>
        /// Combined Bending and Compression EN 1995-1 §6.2.4 - Eq(6.19) + Eq(6.20)
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="Sig0_c_d">Design compressive stress</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        [Description("Combined Bending and Compression EN 1995-1 §6.2.4 - Eq(6.19) + Eq(6.20)")]
        public static double BendingAndCompression(double SigMyd, double SigMzd, double Sig0_c_d, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1, bool FireCheck = false)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");

            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }


            var timber = (IMaterialTimber)material;
            double fmy_k = timber.Fmyk;
            double fmz_k = timber.Fmzk;
            double fc0_k = timber.Fc0k;

            double km = EC5_Factors.Km(crossSection, material);

            double CompressionRatio = Sig0_c_d / (fc0_k * Kmod / Ym * kfi);
            double eq_6_19 = Math.Pow(CompressionRatio, 2) + SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + km * SigMzd / (fmz_k * khz * Kmod / Ym * kfi);
            double eq_6_20 = Math.Pow(CompressionRatio, 2) + km * SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + SigMzd / (fmz_k * khz * Kmod / Ym * kfi);

            return Math.Max(eq_6_19, eq_6_20);
        }


        /// <summary>
        /// Bending and Buckling EN 1995-1 §6.2.4 - Eq(6.23) + Eq(6.24)
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="Sig0_c_d">Design compressive stress</param>
        /// <param name="Leff_Y">Effective Buckling length along Y in mm</param>
        /// <param name="Leff_Z">Effective Buckling Length along Z in mm</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        [Description("Bending and Buckling EN 1995-1 §6.3.2 - Eq(6.23) + Eq(6.24)")]
        public static double BendingAndBuckling(double SigMyd, double SigMzd, double Sig0_c_d, double Leff_Y, double Leff_Z, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1, bool FireCheck = false)
        {

            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }


            var timber = (IMaterialTimber)material;
            double fmy_k = timber.Fmyk;
            double fmz_k = timber.Fmzk;
            double fc0_k = timber.Fc0k;

            double km = EC5_Factors.Km(crossSection, material);
            List<double> Kc = EC5_Factors.Kc(crossSection, timber, Leff_Y, Leff_Z, FireCheck);
            double CompressionRatio = Sig0_c_d / (fc0_k * Kmod / Ym * kfi);

            double eq_6_23 = CompressionRatio / Kc[0] + SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + km * SigMzd / (fmz_k * khz * Kmod / Ym * kfi);
            double eq_6_24 = CompressionRatio / Kc[1] + km * SigMyd / (fmy_k * khy * Kmod / Ym * kfi) + SigMzd / (fmz_k * khz * Kmod / Ym * kfi);

            return Math.Max(eq_6_23, eq_6_24);

        }


        /// <summary>
        /// Lateral Torsional buckling according to DIN EN 1995-1 §6.3.3 Eq(6.33) + Eq(6.35) + Eq(NA.58) + Eq(NA.59) 
        /// </summary>
        /// <param name="SigMyd">Design bending stress in cross section Y axis</param>
        /// <param name="SigMzd">Design bending stress in cross section Z axis</param>
        /// <param name="Sig0_c_d">Design compressive stress</param>
        /// <param name="Leff_Y">Effective Buckling length along Y in mm</param>
        /// <param name="Leff_Z">Effective Buckling Length along Z in mm</param>
        /// <param name="Leff_LTB">Effective Lateral Buckling length in mm</param>
        /// <param name="crossSection">Cross Section Object</param>
        /// <param name="material">Material Object</param>
        /// <param name="Kmod">modification factor</param>
        /// <param name="Ym">Material Safety factor</param>
        /// <param name="khy">Size Factor for Cross section in Y axis</param>
        /// <param name="khz">Size Factor for Cross section in Y axis</param>
        /// <returns>Return the largest value of the 4 equations considered</returns>
        /// <exception cref="Exception"></exception>
        [Description("Lateral Torsional buckling according to DIN EN 1995-1 §6.3.3 Eq(6.33) + Eq(6.35) + Eq(NA.60) + Eq(NA.61)")]
        public static double LateralTorsionalBuckling(double SigMyd, double SigMzd, double Sig0_c_d, double Leff_Y, double Leff_Z, double Leff_LTB, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy = 1, double khz = 1, bool FireCheck = false)
        {

            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");


            //Fire factors
            double kfi = 1;
            if (FireCheck)
            {
                kfi = EC5_Factors.Kfi((IMaterialTimber)material);
                Kmod = 1;
                Ym = 1;
            }

            var timber = (IMaterialTimber)material;
            double fmy_k = timber.Fmyk;
            double fmz_k = timber.Fmzk;
            double fc0_k = timber.Fc0k;


            double kcrit = EC5_Factors.Kcrit(material, crossSection, Leff_LTB, FireCheck);
            List<double> kc = EC5_Factors.Kc(crossSection, material, Leff_Y, Leff_Z, FireCheck);

            double Eq6_33 = 0;
            double Eq6_35 = 0;
            double EqNA_60 = 0;
            double EqNA_61 = 0;


            //The decision has been taken to disregard the clause specifying that Eq(6.33) should be considered only when a moment My acts alone without any accompanying force
            //To our appreciation, this could lead to unsafe designs as a small compressive forces could eventually result in a much lower utilisation ratio


                Eq6_33 = SigMyd / (kcrit * fmy_k * Kmod / Ym * kfi);
                Eq6_35 = Math.Pow(SigMyd / (kcrit * fmy_k * Kmod / Ym * kfi), 2) + Sig0_c_d / (kc[1] * fc0_k * Kmod / Ym * kfi);

                var cs = (CrossSectionRectangular)crossSection;

                //According to DIN EN 1995 §NA.7
                if (cs.H / cs.B >= 4 || (SigMyd > 0 && SigMzd > 0))
                {
                    EqNA_60 = Sig0_c_d / (kc[0] * fc0_k * Kmod / Ym * kfi) + SigMyd / (kcrit * fmy_k * Kmod / Ym * kfi) + Math.Pow(SigMzd / (fmz_k * Kmod / Ym * kfi), 2);

                    EqNA_61 = Sig0_c_d / (kc[1] * fc0_k * Kmod / Ym * kfi) + Math.Pow(SigMyd / (kcrit * fmy_k * Kmod / Ym * kfi), 2) + SigMzd / (fmz_k * Kmod / Ym * kfi);
                }
            

            List<double> results = new List<double>() { Eq6_33, Eq6_35, EqNA_60, EqNA_61 }.OrderByDescending(p => p).ToList();

            return results[0];
        }



    }
}
