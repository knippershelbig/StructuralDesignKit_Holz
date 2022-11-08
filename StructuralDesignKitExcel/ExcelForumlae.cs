﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;
using Excel = Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Windows;
using System.Runtime.InteropServices;
using System.Runtime.ConstrainedExecution;
using System.Diagnostics.Eventing.Reader;

namespace StructuralDesignKitExcel
{




    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            try
            {
                ExcelDna.IntelliSense.IntelliSenseServer.Install();
            }

            catch (Exception e)
            {

            }

        }


        public void AutoClose()
        {
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();
        }




    }

    public static class ExcelFormulae
    {



        #region Eurocode 5 Factors

        //-------------------------------------------
        //Kmod
        //-------------------------------------------
        [ExcelFunction(Description = "Kmod is a modification factor taking into account the effect of the duration of load and moisture content. Values according to EN 1995-1-1:2004 - Table 3.1",
         Name = "SDK.Factors.Kmod",
         IsHidden = false,
         Category = "SDK.EC5_Factors")]
        public static double Kmod([ExcelArgument(Description = "Timber grade")] string TimberGrade, [ExcelArgument(Description = "Service class (SC1,SC2 or SC3)")] string serviceClass,
            [ExcelArgument(Description = "load Duration" +
            "(Permanent - LongTerm - MediumTerm - ShortTerm - Instantaneous - ShortTerm_Instantaneous")] string loadDuration)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            var SC = EC5_Utilities.GetServiceClass(serviceClass);
            var LC = EC5_Utilities.GetLoadDuration(loadDuration);
            return EC5_Factors.Kmod(mat, SC, LC);
        }


        //-------------------------------------------
        //Ym
        //-------------------------------------------
        [ExcelFunction(Description = "Ym represents the material's safety factor -  Values according to DIN EN 1995-1-1",
            Name = "SDK.Factors.Ym",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Ym([ExcelArgument(Description = "Timber grade")] string TimberGrade)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            return EC5_Factors.Ym(mat);
        }


        //-------------------------------------------
        //Kdef
        //-------------------------------------------
        [ExcelFunction(Description = "Kdef is the Deformation factor to take into account the long time creep behaviour depending on the service class and the timber type - Values according to EN 1995-1-1:2004 - Table 3.2",
             Name = "SDK.Factors.Kdef",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double Kdef([ExcelArgument(Description = "Timber grade")] string TimberGrade, [ExcelArgument(Description = "Service class (SC1,SC2 or SC3)")] string serviceClass)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            var SC = EC5_Utilities.GetServiceClass(serviceClass);

            return EC5_Factors.Kdef(mat, SC);
        }


        //-------------------------------------------
        //Kh_Bending
        //-------------------------------------------
        [ExcelFunction(Description = "The size factor kh considers the inhomogeneities and other deviations from an ideal orthotropic material",
             Name = "SDK.Factors.Kh_Bending",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double KhBending([ExcelArgument(Description = "Timber grade")] string TimberGrade, [ExcelArgument(Description = "Bending height in [mm]")] double height)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            return EC5_Factors.Kh_Bending(mat, height);
        }


        //-------------------------------------------
        //Kh_Tension
        //-------------------------------------------
        [ExcelFunction(Description = "The size factor kh considers the inhomogeneities and other deviations from an ideal orthotropic material",
             Name = "SDK.Factors.Kh_Tension",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double KhTension([ExcelArgument(Description = "Timber grade")] string TimberGrade, [ExcelArgument(Description = "beam height in [mm]")] double height)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            return EC5_Factors.Kh_Tension(mat, height);
        }


        //-------------------------------------------
        //Kl_LVL
        //-------------------------------------------
        [ExcelFunction(Description = "Size factor for LVL(or Baubuche) members submited to tensile force - according to EN 1995-1-1:2004 - Eq(3.4) + Kerto Product Certificate + Baubuche Design Assistance Guide",
             Name = "SDK.Factors.Kl_LVL",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double Kl_LVL([ExcelArgument(Description = "Timber grade")] string TimberGrade, [ExcelArgument(Description = "member length subjected to tension in [mm]")] double length)
        {
            var mat = ExcelHelpers.GetTimberMaterial(TimberGrade).Type;
            return EC5_Factors.Kl_LVL(mat, length);
        }


        //-------------------------------------------
        //Kcr
        //-------------------------------------------
        [ExcelFunction(Description = "Computes the Crack factor for shear resistance Kcr - According to DIN EN 1995-1 NA  §6.1.7(2)",
               Name = "SDK.Factors.Kcr",
               IsHidden = false,
               Category = "SDK.EC5_Factors")]
        public static double Kcr([ExcelArgument(Description = "Timber grade")] string TimberGrade)
        {
            return EC5_Factors.Kcr(ExcelHelpers.GetTimberMaterial(TimberGrade));
        }


        //-------------------------------------------
        //KShape
        //-------------------------------------------
        [ExcelFunction(Description = "Factor depending on the shape of the cross-section (and Material for Baubuche) for Torsion check - According to EN 1995-1 Eq(6.15)",
             Name = "SDK.Factors.KShape",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double KShape([ExcelArgument(Description = "Cross section definition (i.e i.e: CS_R_100x200_GL24h)")] string CrossSection)
        {
            double kshape = 0;
            try
            {
                var CS = ExcelHelpers.CreateRectangularCrossSection(CrossSection);
                kshape = EC5_Factors.KShape(CS, CS.Material);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kshape;
        }


        //-------------------------------------------
        //Km
        //-------------------------------------------
        [ExcelFunction(Description = "Factor considering re-distribution of bending stresses in a cross-section - According to EN 1995-1 §6.1.6(2)",
             Name = "SDK.Factors.Km",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double Km([ExcelArgument(Description = "Cross section definition (i.e i.e: CS_R_100x200_GL24h)")] string CrossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(CrossSection);
            return EC5_Factors.Km(CS, CS.Material);
        }


        //-------------------------------------------
        //Kc
        //-------------------------------------------
        [ExcelFunction(Description = "Computes and returns the buckling instability factors kcy and kcz as a list of doubles - According to EN 1995-1 Eq(6.27) + Eq(6.28)",
             Name = "SDK.Factors.Kc",
             IsHidden = false,
             Category = "SDK.EC5_Factors")]
        public static double Kc(
            [ExcelArgument(Description = "Cross section definition (i.e i.e: CS_R_100x200_GL24h)")] string CrossSection,
            [ExcelArgument(Description = "Buckling Length relative to Y in [mm]")] double BucklingLength_Y,
            [ExcelArgument(Description = "Buckling Length relative to Z in [mm]")] double BucklingLength_Z,
            [ExcelArgument(Description = "Kcy = 0, Kcz = 1")] int axis)
        {
            double Kc = 0;

            try
            {
                if (axis != 0 && axis != 1) throw new Exception("Axis should be either 0 or 1");
                var CS = ExcelHelpers.CreateRectangularCrossSection(CrossSection);
                Kc = EC5_Factors.Kc(CS, CS.Material, BucklingLength_Y, BucklingLength_Z)[axis];

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return Kc;
        }


        //-------------------------------------------
        //Kcrit
        //-------------------------------------------
        [ExcelFunction(Description = "factor which takes into account the reduced bending strength due to lateral buckling according to EN 1995-1 Eq(6.34)",
            Name = "SDK.Factors.Kcrit",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kcrit(
            [ExcelArgument(Description = "Cross section definition (i.e i.e: CS_R_100x200_GL24h)")] string CrossSection,
            [ExcelArgument(Description = "Lateral Buckling Length in [mm]")] double LatBucklingLength)
        {
            double kcrit = 0;

            try
            {
                var CS = ExcelHelpers.CreateRectangularCrossSection(CrossSection);
                kcrit = EC5_Factors.Kcrit(CS.Material, CS, LatBucklingLength);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kcrit;
        }


        //-------------------------------------------
        //Kc90
        //-------------------------------------------
        [ExcelFunction(Description = "kc,90 takes into consideration the type of effect, the splitting risk and the extent of the deformation",
            Name = "SDK.Factors.Kc90",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kc90(
            [ExcelArgument(Description = "Timber grade")] string TimberGrade,
            [ExcelArgument(Description = "Support Type: 0 for continuous support; 1 for ponctual contact according to EN 1995-1 §6.1.5")] int SupportType)
        {
            double kc90 = 0;

            try
            {
                if (SupportType != 0 && SupportType != 1) throw new Exception("Support type should be 0 or 1");
                var mat = ExcelHelpers.GetTimberMaterial(TimberGrade);
                kc90 = EC5_Factors.Kc90(mat, SupportType);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kc90;
        }


        //-------------------------------------------
        //KmAlphaCompression
        //-------------------------------------------
        [ExcelFunction(Description = "Km,α is a reduction factor taking into account the increased stress due to tapper edges. The factor differs for edges in tension or compression" +
            "depending on the bending orientation - EN 1995-1 §6.4.2 eq (6.39)",
            Name = "SDK.Factors.KmAlphaComp",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double KmAlphaCompression(
            [ExcelArgument(Description = "Timber grade")] string TimberGrade,
            [ExcelArgument(Description = "Cut angle relative to the grain in [° Degree]")] double angle)
        {
            double kmalpha = 0;

            try
            {
                var mat = ExcelHelpers.GetTimberMaterial(TimberGrade);
                kmalpha = EC5_Factors.Km_Alpha_Compression(mat, angle);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kmalpha;
        }


        //-------------------------------------------
        //KmAlphaTension
        //-------------------------------------------
        [ExcelFunction(Description = "Km,α is a reduction factor taking into account the increased stress due to tapper edges. The factor differs for edges in tension or compression" +
            "depending on the bending orientation - EN 1995-1 §6.4.2 eq (6.39)",
            Name = "SDK.Factors.KmAlphaTension",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double KmAlphaTension(
            [ExcelArgument(Description = "Timber grade")] string TimberGrade,
            [ExcelArgument(Description = "Cut angle relative to the grain in [° Degree]")] double angle)
        {
            double kmalpha = 0;

            try
            {
                var mat = ExcelHelpers.GetTimberMaterial(TimberGrade);
                kmalpha = EC5_Factors.Km_Alpha_Tension(mat, angle);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kmalpha;
        }


        //-------------------------------------------
        //Kr
        //-------------------------------------------
        [ExcelFunction(Description = "Factor taking into consideration the stresses generated prior to bonding due to the bending of individual" +
            "lamellae for curved beams with small radii of curvature - EN 1995-1 §6.4.3 eq (6.49)",
            Name = "SDK.Factors.Kr",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kr(
            [ExcelArgument(Description = "Internal beam radius in [mm]")] double internalRadius,
            [ExcelArgument(Description = "glulam lamellae thickness in [mm]")] double lamellaThickness)
        {
            double kr = 0;

            try
            {
                kr = EC5_Factors.Kr(internalRadius, lamellaThickness);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kr;
        }


        //-------------------------------------------
        //Kl
        //-------------------------------------------
        [ExcelFunction(Description = "Kl is a bending amplification factor taking into account the beam curvature,cut angle and height in the apex area - EN 1995-1 §6.4.3 eq (6.43)",
            Name = "SDK.Factors.Kl",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kl(
            [ExcelArgument(Description = "Beam height at apex in [mm]")] double heightApex,
            [ExcelArgument(Description = "cut angle at apex in degree")] double angleApex,
            [ExcelArgument(Description = "beam internal radius in [mm]")] double internalRadius)
        {
            double kl = 0;

            try
            {
                kl = EC5_Factors.Kl(heightApex, angleApex, internalRadius);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kl;
        }


        //-------------------------------------------
        //Kvol
        //-------------------------------------------
        [ExcelFunction(Description = "Volume factor taking into consideration the influence of volume on tensile strength perpendicular to the grain - EN 1995-1 §6.4.3 eq (6.51)",
            Name = "SDK.Factors.Kvol",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kvol(
            [ExcelArgument(Description = "Material Object")] string material,
            [ExcelArgument(Description = "stressed volume of the apex zone in [m³]")] double Vstressed,
            [ExcelArgument(Description = "Total beam volume in [m³]")] double Vtot)

        {
            double kvol = 0;

            try
            {
                var mat = ExcelHelpers.GetTimberMaterial(material);
                kvol = EC5_Factors.Kvol(mat, Vstressed, Vtot);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kvol;
        }


        //-------------------------------------------
        //Kdis
        //-------------------------------------------
        [ExcelFunction(Description = "Factor taking into consideration the influence of stress distribution - EN 1995-1 §6.4.3 eq (6.52)",
            Name = "SDK.Factors.Kdis",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kdis(
            [ExcelArgument(Description = "0->double tapered and curved beams | 1-> for pitched cambered beams")] int beamType)
        {
            double kdis = 0;

            try
            {
                if (beamType != 0 && beamType != 1) throw new Exception("beamType should be either 0 or 1");
                kdis = EC5_Factors.Kdis(beamType);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kdis;
        }


        //-------------------------------------------
        //Kp
        //-------------------------------------------
        [ExcelFunction(Description = "Factor for the verification of tension perpendicular to the grain at apex - EN 1995-1 §6.4.3 eq (6.56)",
            Name = "SDK.Factors.Kp",
            IsHidden = false,
            Category = "SDK.EC5_Factors")]
        public static double Kp([ExcelArgument(Description = "Beam height at apex in [mm]")] double heightApex, [ExcelArgument(Description = "cutting angle at apex in degree")] double angleApex,
            [ExcelArgument(Description = "beam internal radius in [mm]")] double internalRadius)
        {
            double kp = 0;

            try
            {

                kp = EC5_Factors.Kp(heightApex, angleApex, internalRadius);

            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return kp;
        }




        #endregion


        #region Eucocode 5 Cross-section checks


        //-------------------------------------------
        //Tension parallel to the grain §6.1.2
        //-------------------------------------------
        [ExcelFunction(Description = "Tension parallel to the grain §6.1.2",
            Name = "SDK.CrossSectionChecks.TensionParallelToGrain_6.1.2",
            IsHidden = false,
            Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_TensionParallelToGrain(
            [ExcelArgument(Description = "Design tensile stress")] double Sig0_t_d,
            [ExcelArgument(Description = "Timber grade")] string timberGrade,
            [ExcelArgument(Description = "Kmod factor")] double Kmod,
            [ExcelArgument(Description = "safety factor Ym")] double Ym,
            [ExcelArgument(Description = "Size Factor for Cross section")] double Kh,
            [ExcelArgument(Description = "Mofification factor for LVL member length")] double Kl_LVL = 1)
        {
            var mat = ExcelHelpers.GetTimberMaterial(timberGrade);

            return EC5_CrossSectionCheck.TensionParallelToGrain(Sig0_t_d, mat, Kmod, Ym, Kh, Kl_LVL);
        }



        //-------------------------------------------
        //Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)
        //-------------------------------------------
        [ExcelFunction(Description = "Compression parallel to the grain EN 1995-1 §6.1.4 - Eq(6.2)",
            Name = "SDK.CrossSectionChecks.CompressionParallelToGrain_6.1.4",
            IsHidden = false,
            Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_CompressionParallelToGrain(
            [ExcelArgument(Description = "Design compressive stress")] double Sig0_c_d,
            [ExcelArgument(Description = "Timber grade")] string timberGrade,
            [ExcelArgument(Description = "Kmod factor")] double Kmod,
            [ExcelArgument(Description = "safety factor Ym")] double Ym)
        {
            var mat = ExcelHelpers.GetTimberMaterial(timberGrade);
            return EC5_CrossSectionCheck.CompressionParallelToGrain(Sig0_c_d, mat, Kmod, Ym);
        }


        //-------------------------------------------
        //Compression stresses at an angle to the grain EN 1995-1 §6.2.2 - Eq(6.16)
        //-------------------------------------------
        [ExcelFunction(Description = "Compression stresses at an angle to the grain EN 1995-1 §6.2.2 - Eq(6.16)",
            Name = "SDK.CrossSectionChecks.CompressionAtAnAngleToGrain_6.2.2",
            IsHidden = false,
            Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_CompressionAtAnAngleToGrain(
            [ExcelArgument(Description = "Design compressive stress")] double SigAlpha_c_d,
            [ExcelArgument(Description = "Timber grade")] string timberGrade,
            [ExcelArgument(Description = "stress angle to the grain in Degree")] double angleToGrain,
            [ExcelArgument(Description = "Kmod factor")] double Kmod,
            [ExcelArgument(Description = "safety factor Ym")] double Ym,
            [ExcelArgument(Description = "factor taking into account the effect of stresses perpendicular to the grain")] double kc90)
        {
            var mat = ExcelHelpers.GetTimberMaterial(timberGrade);
            return EC5_CrossSectionCheck.CompressionAtAnAngleToGrain(SigAlpha_c_d, angleToGrain, mat, Kmod, Ym, kc90 = 1);
        }


        //-------------------------------------------
        //Bending EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12)
        //-------------------------------------------
        [ExcelFunction(Description = "Bending EN 1995-1 §6.1.6 - Eq(6.11) + Eq(6.12)",
            Name = "SDK.CrossSectionChecks.Bending_6.1.6",
            IsHidden = false,
            Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_Bending(
            [ExcelArgument(Description = "Design bending stress in cross section Y axis")] double SigMyd,
            [ExcelArgument(Description = "Design bending stress in cross section Z axis")] double SigMzd,
            [ExcelArgument(Description = "Cross section to check")] string crossSection,
            [ExcelArgument(Description = "Kmod factor")] double Kmod,
            [ExcelArgument(Description = "safety factor Ym")] double Ym,
            [ExcelArgument(Description = "Size Factor for Cross section in Y axis")] double khy = 1,
            [ExcelArgument(Description = "Size Factor for Cross section in Z axis")] double khz = 1)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.Bending(SigMyd, SigMzd, CS, mat, Kmod, Ym, khy, khz);
        }


        //-------------------------------------------
        //Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)
        //-------------------------------------------
        [ExcelFunction(Description = "Shear DIN EN 1995-1 +NA §6.1.7 - Eq(6.13) + Eq(6.13a) + Eq(NA.54)",
            Name = "SDK.CrossSectionChecks.Shear_6.1.7",
            IsHidden = false,
            Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_Shear(
            [ExcelArgument(Description = "Design shear stress on Y")] double TauYd,
            [ExcelArgument(Description = "Design shear stress on Z")] double TauZd,
            [ExcelArgument(Description = "Timber grade")] string timberGrade,
            [ExcelArgument(Description = "Kmod factor")] double Kmod,
            [ExcelArgument(Description = "safety factor Ym")] double Ym)
        {
            var mat = ExcelHelpers.GetTimberMaterial(timberGrade);
            return EC5_CrossSectionCheck.Shear(TauYd, TauZd, mat, Kmod, Ym);
        }


        //-------------------------------------------
        //Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)
        //-------------------------------------------
        [ExcelFunction(Description = "Torsion DIN EN 1995-1 +NA §6.1.8 - Eq(6.15) + Eq(NA.55)",
             Name = "SDK.CrossSectionChecks.Torsion_6.1.8",
             IsHidden = false,
             Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_Torsion(
             [ExcelArgument(Description = "Design torsion shear stress")] double TauTorsion,
             [ExcelArgument(Description = "Design shear stress on Y")] double TauYd,
             [ExcelArgument(Description = "Design shear stress on Z")] double TauZd,
             [ExcelArgument(Description = "Cross section")] string crossSection,
             [ExcelArgument(Description = "Kmod factor")] double Kmod,
             [ExcelArgument(Description = "safety factor Ym")] double Ym)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.Torsion(TauTorsion, TauYd, TauZd, CS, mat, Kmod, Ym);
        }


        //-------------------------------------------
        //Bending and tension EN 1995-1 +NA §6.2.3 Eq(6.17) + Eq(6.18)
        //-------------------------------------------
        [ExcelFunction(Description = "Bending and tension EN 1995-1 +NA §6.2.3 Eq(6.17) + Eq(6.18)",
             Name = "SDK.CrossSectionChecks.BendingAndTension_6.2.3",
             IsHidden = false,
             Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_BendingAndTension(
             [ExcelArgument(Description = "Design bending stress in cross section Y axis")] double SigMyd,
             [ExcelArgument(Description = "Design bending stress in cross section Z axis")] double SigMzd,
             [ExcelArgument(Description = "Design tensile stress")] double Sig0_t_d,
             [ExcelArgument(Description = "Cross section")] string crossSection,
             [ExcelArgument(Description = "Kmod factor")] double Kmod,
             [ExcelArgument(Description = "safety factor Ym")] double Ym,
             [ExcelArgument(Description = "Size Factor for Cross section in Y axis")] double khy = 1,
             [ExcelArgument(Description = "Size Factor for Cross section in Z axis")] double khz = 1,
             [ExcelArgument(Description = "Size Factor for Cross section in tension")] double Kh_Tension = 1,
             [ExcelArgument(Description = "Mofification factor for member Length")] double Kl_LVL = 1)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.BendingAndTension(SigMyd, SigMzd, Sig0_t_d, CS, mat, Kmod, Ym, khy, khz, Kh_Tension, Kl_LVL);
        }


        //-------------------------------------------
        //Combined Bending and Compression EN 1995-1 §6.2.4 - Eq(6.19) + Eq(6.20)
        //-------------------------------------------
        [ExcelFunction(Description = "Combined Bending and Compression EN 1995-1 §6.2.4 - Eq(6.19) + Eq(6.20)",
             Name = "SDK.CrossSectionChecks.BendingAndCompression_6.2.4",
             IsHidden = false,
             Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_BendingAndCompression(
             [ExcelArgument(Description = "Design bending stress in cross section Y axis")] double SigMyd,
             [ExcelArgument(Description = "Design bending stress in cross section Z axis")] double SigMzd,
             [ExcelArgument(Description = "Design compressive stress")] double Sig0_c_d,
             [ExcelArgument(Description = "Cross section")] string crossSection,
             [ExcelArgument(Description = "Kmod factor")] double Kmod,
             [ExcelArgument(Description = "safety factor Ym")] double Ym,
             [ExcelArgument(Description = "Size Factor for Cross section in Y axis")] double khy = 1,
             [ExcelArgument(Description = "Size Factor for Cross section in Z axis")] double khz = 1)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.BendingAndCompression(SigMyd, SigMzd, Sig0_c_d, CS, mat, Kmod, Ym, khy, khz); ;
        }


        //-------------------------------------------
        //Bending and Buckling EN 1995-1 §6.3.2 - Eq(6.23) + Eq(6.24)
        //-------------------------------------------
        [ExcelFunction(Description = "Bending and Buckling EN 1995-1 §6.3.2 - Eq(6.23) + Eq(6.24)",
             Name = "SDK.CrossSectionChecks.BendingAndBuckling_6.3.2",
             IsHidden = false,
             Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_BendingAndBuckling(
             [ExcelArgument(Description = "Design bending stress in cross section Y axis")] double SigMyd,
             [ExcelArgument(Description = "Design bending stress in cross section Z axis")] double SigMzd,
             [ExcelArgument(Description = "Design compressive stress")] double Sig0_c_d,
             [ExcelArgument(Description = "Effective Buckling length along Y in mm")] double Leff_Y,
             [ExcelArgument(Description = "Effective Buckling length along Z in mm")] double Leff_Z,
             [ExcelArgument(Description = "Cross section")] string crossSection,
             [ExcelArgument(Description = "Kmod factor")] double Kmod,
             [ExcelArgument(Description = "safety factor Ym")] double Ym,
             [ExcelArgument(Description = "Size Factor for Cross section in Y axis")] double khy = 1,
             [ExcelArgument(Description = "Size Factor for Cross section in Z axis")] double khz = 1)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.BendingAndBuckling(SigMyd, SigMzd, Sig0_c_d, Leff_Y, Leff_Z, CS, mat, Kmod, Ym, khy, khz);
        }


        //-------------------------------------------
        //Lateral Torsional buckling according to DIN EN 1995-1 §6.3.3 Eq(6.33) + Eq(6.35) + Eq(NA.58) + Eq(NA.59)
        //-------------------------------------------
        [ExcelFunction(Description = "Lateral Torsional buckling according to DIN EN 1995-1 §6.3.3 Eq(6.33) + Eq(6.35) + Eq(NA.58) + Eq(NA.59)",
             Name = "SDK.CrossSectionChecks.LateralTorsionalBuckling_6.3.3",
             IsHidden = false,
             Category = "SDK.EC5_CrossSection_Checks")]
        public static double CS_Check_LateralTorsionalBuckling(
             [ExcelArgument(Description = "Design bending stress in cross section Y axis")] double SigMyd,
             [ExcelArgument(Description = "Design bending stress in cross section Z axis")] double SigMzd,
             [ExcelArgument(Description = "Design compressive stress")] double Sig0_c_d,
             [ExcelArgument(Description = "Effective Buckling length along Y in mm")] double Leff_Y,
             [ExcelArgument(Description = "Effective Buckling length along Z in mm")] double Leff_Z,
             [ExcelArgument(Description = "Effective Lateral Buckling length in mm")] double Leff_LTB,
             [ExcelArgument(Description = "Cross section")] string crossSection,
             [ExcelArgument(Description = "Kmod factor")] double Kmod,
             [ExcelArgument(Description = "safety factor Ym")] double Ym,
             [ExcelArgument(Description = "Size Factor for Cross section in Y axis")] double khy = 1,
             [ExcelArgument(Description = "Size Factor for Cross section in Z axis")] double khz = 1)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            var mat = CS.Material;
            return EC5_CrossSectionCheck.LateralTorsionalBuckling(SigMyd, SigMzd, Sig0_c_d, Leff_Y, Leff_Z, Leff_LTB, CS, mat, Kmod, Ym, khy, khz);
        }

        #endregion



        #region Material
        //-------------------------------------------
        //Material Property
        //-------------------------------------------
        [ExcelFunction(Description = "Retrieve the material property, given a cross section and the property name",
         Name = "SDK.Material.Property",
         IsHidden = false,
         Category = "SDK.EC5_Material")]
        public static double MaterialProperty([ExcelArgument(Description = "Cross Section")] string CrossSection, [ExcelArgument(Description = "Property name")] string propertyName)
        {
            try
            {
                var CS = ExcelHelpers.CreateRectangularCrossSection(CrossSection);
                double returnValue = 0;

                //Get Timber properties
                var properties = typeof(StructuralDesignKitLibrary.Materials.IMaterialTimber).GetProperties().ToList();
                //return car.GetType().GetProperty(propertyName).GetValue(car, null);
                if (properties.Where(p => p.Name == propertyName).ToList().Count > 0) //The property given exists
                {
                    if (propertyName != "Type")
                    {
                        var timberMat = CS.Material as StructuralDesignKitLibrary.Materials.IMaterialTimber;
                        returnValue = (double)timberMat.GetType().GetProperty(propertyName).GetValue(timberMat);
                    }

                }
                else throw new Exception("Material property unknow");

                return returnValue;
            }
            catch (Exception)
            {
                return -1;
            }
         
        }



        #endregion



        #region Compute Stresses



        //-------------------------------------------
        //BendingY
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the bending stresses on the Y axis for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.BendingY",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeBendingStressY(
           [ExcelArgument(Description = "Bending moment My in [KN.m]")] double My,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeStressBendingY(My);
        }

        //-------------------------------------------
        //BendingZ
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the bending stresses on the Z axis for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.BendingZ",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeBendingStressZ(
           [ExcelArgument(Description = "Bending moment Mz in [KN.m]")] double Mz,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeStressBendingZ(Mz);
        }

        //-------------------------------------------
        //Normal stress
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the normal stress for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.NormalForce",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeNormalStress(
           [ExcelArgument(Description = "Normal Force in [KN]")] double N,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeNormalStress(N);
        }

        //-------------------------------------------
        //Shear Y
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the shear stress on the Z axis for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.ShearY",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeShearStressY(
           [ExcelArgument(Description = "Shear Force on Y axis in [KN]")] double Vy,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeShearY(Vy);
        }


        //-------------------------------------------
        //Shear Z
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the shear stress on the Z axis for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.ShearZ",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeShearStressZ(
           [ExcelArgument(Description = "Shear Force on Z axis in [KN]")] double Vz,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeShearZ(Vz);
        }


        //-------------------------------------------
        //Torsion X
        //-------------------------------------------
        [ExcelFunction(Description = "Compute the shear stress due to torsion on the X axis for a rectangular cross section - Result in [N/mm²]",
           Name = "SDK.CrossSection_StressCompute.TorsionShear",
           IsHidden = false,
           Category = "SDK.CrossSection_StressCompute")]
        public static double CS_ComputeTorsion(
           [ExcelArgument(Description = "Torsion moment Mx [KN.m]")] double Mx,
           [ExcelArgument(Description = "Cross section")] string crossSection)
        {
            var CS = ExcelHelpers.CreateRectangularCrossSection(crossSection);
            return CS.ComputeTorsion(Mx);
        }

        #endregion



        #region utilities
        [ExcelFunction(Description = "Create a cross section tag",
            Name = "SDK.Material.CreateRectangularCrossSection",
            IsHidden = false,
            Category = "SDK.Utilities")]
        public static string CreateCrossSection([ExcelArgument(Description = "width")] double b, [ExcelArgument(Description = "height")] double h, string material)
        {
            return ExcelHelpers.CreateRectangularCrossSection(b, h, ExcelHelpers.GetTimberMaterial(material));
        }

        #endregion


        //material encryypt and decript

        //Check all at once given a load torsor

        //Create drop down list with materials

        //Create a list of all material properties in excel (kind of deconstruct material) - Function

        //Create a material which can be used following the IMaterial Timber Interface

        //Expose drop down Material type (baubuche, ...)

        //Generate functions (+ Buttons in Ribbon) to display factors + table of factors

        //Generate functions with boiling plates for the different checks and full checks - Highlight cell to modify

        //Generate material tables

        //Dropdown enum (material type, load duration, service class, ...)

        //Add optimisation (Cross section / material for given check)

        //LTB according to EC3





        #region Garbage Collector

        //public static StructuralDesignKitLibrary.Materials.IMaterialTimber mat()
        //{
        //    return new MaterialTimberBaubuche()
        //}





        //[ExcelFunction("Returns the answer")]
        //public static object MyFunction([ExcelArgument("The unimportant input")] object input)
        //{
        //    return 42;
        //}





        //[ExcelFunction(Description = "Find a material based on a string",
        //    Name = "SDK.Material.GetMaterial",
        //    IsHidden = false,
        //    Category = "SDK.EC5_Factors")]
        //public static double buckling()
        //{
        //    var Kcs = EC5_Factors.Kc()
        //}






        //[ExcelFunction(Description = "TestFunction",
        //Name = "SDK.Material.Test",
        //IsHidden = false,
        //Category = "SDK.EC5_Materials")]
        //public static string CrossSection(string CrossSectionTag)
        //{
        //    CrossSectionRectangular CS = ExcelHelpers.CreateCrossSection(CrossSectionTag);
        //    return String.Format("{0}x{1}mm", CS.B.ToString(), CS.H.ToString());

        //}






        //public static class MyFunctions
        //{

        //    [ExcelFunction(Description = "My first .NET function")]
        //    public static string SayHello(string name)
        //    {
        //        return "Hello " + name;
        //    }

        //    [ExcelFunction(Description = "Return a SDK Material")]
        //    public static Object material(string materialName)
        //    {
        //        var mat = new MaterialTimberGlulam(materialName);
        //        var guid = Guid.NewGuid();
        //        return mat.Ft0k;
        //    }

        //    [ExcelFunction(Description = "Return a FCK from a material")]
        //    public static double fc0k(object material)
        //    {
        //        var mat = (IMaterialTimber)material;
        //        return mat.Fc0k;
        //    }

        //    //        // Get the type corresponding to the class MyClass.
        //    //Type myType = typeof(MyClass1);
        //    //// Get the object of the Guid.
        //    //Guid myGuid = (Guid)myType.GUID;
        //    //Console.WriteLine("The name of the class is "+myType.ToString());
        //    //Console.WriteLine("The ClassId of MyClass is "+myType.GUID);	


        //    [ExcelFunction("Returns the answer")]
        //    public static object MyFunction([ExcelArgument("The unimportant input")] object input)
        //    {
        //        return 42;
        //    }


        //}

        //public class testclass : XlCall
        //{
        //    //test function objects
        //    [ExcelFunction(Description = "TestFunction returning objects to Excel",
        //     Name = "SDK.testObject",
        //     IsHidden = false,
        //     Category = "SDK.Test")]
        //    public object TestObject(
        //         [ExcelArgument(Name = "SourceData", Description = "The range of cells to be transposed.")] object oSource,
        //         [ExcelArgument(Name = "ByRow", Description = "Optional flag to force transposing vertically insted of the horizontal default.")][Optional] bool bByRow
        //                                    )

        //    {
        //        var oCaller = Excel(xlfCaller) as ExcelReference;
        //        return oCaller.GetValue();
        //    }
        //}

        #endregion



    }

}






