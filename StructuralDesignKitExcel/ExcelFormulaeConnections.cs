using ExcelDna.Integration;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.Connections.Fasteners;
using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using StructuralDesignKitLibrary.Connections.TimberTimberShear;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.Vibrations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace StructuralDesignKitExcel
{
    public static partial class ExcelFormulae
    {
        #region Connections

        [ExcelFunction(Description = "Characteristic yield moment of the fastener in N.mm",
            Name = "SDK.EC5.Connections.Myrk",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double Myrk([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerTag)
        {
            IFastener fastener = ExcelHelpers.GetFastenerFromTag(fastenerTag);
            return fastener.MyRk;
        }


        [ExcelFunction(Description = "Embedment Strength of the fastener in N/mm²",
            Name = "SDK.EC5.Connections.Fhk",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double Fhk([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerTag,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle,
            [ExcelArgument(Description = "Timber type")] string timber)
        {
            IFastener fastener = ExcelHelpers.GetFastenerFromTag(fastenerTag);
            IMaterialTimber timberMaterial = ExcelHelpers.GetTimberMaterialFromTag(timber);
            fastener.ComputeEmbedmentStrength(timberMaterial, angle);
            return fastener.Fhk;
        }


        [ExcelFunction(Description = "Get the percentage of the Johansen Part that the rope effect can contribute to",
            Name = "SDK.EC5.Connections.MaxJohansenPart",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double MaxJohansenPart([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType)
        {

            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, 20, 800);
            return fastener.MaxJohansenPart;
        }


        [ExcelFunction(Description = "K90 is a factor taking into consideration the splitting risk and the degree of compressive deformation - According to EN 1995-1-1 Eq (8.33)",
             Name = "SDK.EC5.Connections.K90",
             IsHidden = false,
             Category = "SDK.EC5.Connections_Utilities")]
        public static double K90(
            [ExcelArgument(Description = "String representing the fastener type (i.e Bolt_D10_Fu800)")] string fastenerTag,
            [ExcelArgument(Description = "String representing the timber grade")] string timber)
        {
            IFastener fastener = ExcelHelpers.GetFastenerFromTag(fastenerTag);
            var timberObj = ExcelHelpers.GetTimberMaterialFromTag(timber);

            var bolt = new FastenerBolt(fastener.Diameter, fastener.Fuk);
            return bolt.ComputeK90(timberObj);
        }


        [ExcelFunction(Description = "Minimum spacing parallel to grain in mm",
            Name = "SDK.EC5.Connections.a1Min",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A1Min([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {

            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);

            fastener.ComputeSpacings(angle);
            return fastener.a1min;
        }


        [ExcelFunction(Description = "Minimum spacing perpendicular to grain in mm",
            Name = "SDK.EC5.Connections.a2Min",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A2Min([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {
            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);
            fastener.ComputeSpacings(angle);
            return fastener.a2min;
        }


        [ExcelFunction(Description = "Minimum spacing to loaded end in mm",
            Name = "SDK.EC5.Connections.a3tMin",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A3tMin([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {
            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);
            fastener.ComputeSpacings(angle);
            return fastener.a3tmin;
        }


        [ExcelFunction(Description = "Minimum spacing to unloaded end in mm",
            Name = "SDK.EC5.Connections.a3cMin",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A3cMin([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {
            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);
            fastener.ComputeSpacings(angle);
            return fastener.a3cmin;
        }


        [ExcelFunction(Description = "Minimum spacing to loaded edge in mm",
            Name = "SDK.EC5.Connections.a4tMin",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A4tMin([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {
            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);
            fastener.ComputeSpacings(angle);
            return fastener.a4tmin;
        }


        [ExcelFunction(Description = "Minimum spacing to unloaded edge in mm",
            Name = "SDK.EC5.Connections.a4cMin",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double A4cMin([ExcelArgument(Description = "String representing the fastener type (i.e \"Bolt\", \"Dowel\", ...")] string fastenerType,
            [ExcelArgument(Description = "Fastener diameter in mm")] double diameter,
            [ExcelArgument(Description = "Load angle to the timber grain")] double angle)
        {
            IFastener fastener = null;
            if (ExcelHelpers.IsFastener(fastenerType)) fastener = ExcelHelpers.GetFastener(fastenerType, diameter, 800);
            fastener.ComputeSpacings(angle);
            return fastener.a4cmin;
        }


        [ExcelFunction(Description = "Compute the effective number of fastener to consider",
            Name = "SDK.EC5.Connections.Neff",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double Neff(
            [ExcelArgument(Description = "String representing the fastener type (i.e Bolt_D10_Fu800)")] string fastenerTag,
            [ExcelArgument(Description = "Number of fasterner in a row")] int n,
            [ExcelArgument(Description = "Spacing alongside the grain a1 in [mm]")] double a1,
            [ExcelArgument(Description = "Load angle to the grain in [degree]")] double angle)
        {
            IFastener fastener = ExcelHelpers.GetFastenerFromTag(fastenerTag);
            return fastener.ComputeEffectiveNumberOfFastener(n, a1, angle);
        }



        /*
         * -------------------------
         * STEEL TO TIMBER CONNECTIONS
         * -------------------------
         */

        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to steel connection with inner plate in [N]. The value is for 2 shear planes",
            Name = "SDK.EC5.Connection.SingleInnerSteelPlate",
            IsHidden = false,
            Category = "SDK.EC5.Connections_SteelTimber")]
        public static double SingleInnerSteelPlate(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            return new SingleInnerSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect).Capacity;
        }


        [ExcelFunction(Description = "Return the shear capacity of a chosen failure mode for a timber to steel connection with inner plate in [N]. The value is for 1 shear plane",
            Name = "SDK.EC5.Connection.SingleInnerSteelPlateFailureMode",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double SingleInnerSteelPlateFailureMode(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect,
            [ExcelArgument(Description = "Failure mode to consider")] string failureMode)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            var connection = new SingleInnerSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect);

            if (connection.FailureModes.Contains(failureMode)) return connection.Capacities[connection.FailureModes.IndexOf(failureMode)];
            else return -1.0;
        }


        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to steel connection with one outer plate in [N]",
            Name = "SDK.EC5.Connection.SingleOuterSteelPlate",
            IsHidden = false,
            Category = "SDK.EC5.Connections_SteelTimber")]
        public static double SingleOuterSteelPlate(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1 to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            return new SingleOuterSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect).Capacity;
        }

        [ExcelFunction(Description = "Return the shear capacity of a chosen failure mode for a timber to steel connection with one outer plate in [N]",
            Name = "SDK.EC5.Connection.SingleOuterSteelPlateFailureMode",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double SingleOuterSteelPlateFailureMode(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1) according to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect,
            [ExcelArgument(Description = "Failure mode to consider")] string failureMode)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            var connection = new SingleOuterSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect);

            if (connection.FailureModes.Contains(failureMode)) return connection.Capacities[connection.FailureModes.IndexOf(failureMode)];
            else return -1.0;
        }



        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to steel connection with two outer plates in [N];The value is for 2 shear planes",
            Name = "SDK.EC5.Connection.DoubleOuterSteelPlate",
            IsHidden = false,
            Category = "SDK.EC5.Connections_SteelTimber")]
        public static double DoubleOuterSteelPlate(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1 to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            return new DoubleOuterSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect).Capacity;
        }

        [ExcelFunction(Description = "Return the shear capacity of a chosen failure mode for a timber to steel connection with one outer plate in [N]; The value is for 1 shear plane",
            Name = "SDK.EC5.Connection.DoubleOuterSteelPlateFailureMode",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double DoubleOuterSteelPlateFailureMode(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "Steel plate thickness in [mm]")] double steelPlateThickness,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle,
            [ExcelArgument(Description = "Timber material used in the connection")] string material,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane (t1) according to 8.2.3 (3))")] double timberThickness,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect,
            [ExcelArgument(Description = "Failure mode to consider")] string failureMode)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var timber = ExcelHelpers.GetTimberMaterialFromTag(material);
            var connection = new DoubleOuterSteelPlate(fastener, steelPlateThickness, angle, timber, timberThickness, ropeEffect);

            if (connection.FailureModes.Contains(failureMode)) return connection.Capacities[connection.FailureModes.IndexOf(failureMode)];
            else return -1.0;
        }



        /*
         * -------------------------
         * TIMBER TO TIMBER CONNECTIONS
         * -------------------------
         */
        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to timber connection in [N]. The value is for 1 shear plane",
            Name = "SDK.EC5.Connection.TimberTimberSingleShear",
            IsHidden = false,
            Category = "SDK.EC5.Connections_TimberTimber")]
        public static double TimberTimberSingleShear(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "First timber material used in the shear plane")] string timber1,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber1,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle1,
            [ExcelArgument(Description = "Second timber material used in the shear plane")] string timber2,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber2,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle2,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var Timber1 = ExcelHelpers.GetTimberMaterialFromTag(timber1);
            var Timber2 = ExcelHelpers.GetTimberMaterialFromTag(timber2);

            return new TimberTimberSingleShear(fastener, Timber1, ThicknessTimber1, angle1, Timber2, ThicknessTimber2, angle2, ropeEffect).Capacity;
        }

        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to timber connection in [N]. The value is for 1 shear plane",
            Name = "SDK.EC5.Connection.TimberTimberSingleShearFailureMode",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double TimberTimberSingleShearFailureMode(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "First timber material used in the shear plane")] string timber1,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber1,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle1,
            [ExcelArgument(Description = "Second timber material used in the shear plane")] string timber2,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber2,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle2,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect,
            [ExcelArgument(Description = "Failure mode to consider")] string failureMode)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var Timber1 = ExcelHelpers.GetTimberMaterialFromTag(timber1);
            var Timber2 = ExcelHelpers.GetTimberMaterialFromTag(timber2);

            var connection = new TimberTimberSingleShear(fastener, Timber1, ThicknessTimber1, angle1, Timber2, ThicknessTimber2, angle2, ropeEffect);

            if (connection.FailureModes.Contains(failureMode)) return connection.Capacities[connection.FailureModes.IndexOf(failureMode)];
            else return -1.0;
        }



        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to timber connection in [N]. The value is for 2 shear plane",
         Name = "SDK.EC5.Connection.TimberTimberDoubleShear",
         IsHidden = false,
         Category = "SDK.EC5.Connections_TimberTimber")]
        public static double TimberTimberDoubleShear(
         [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
         [ExcelArgument(Description = "First timber material used in the shear plane")] string timber1,
         [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber1,
         [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle1,
         [ExcelArgument(Description = "Second timber material used in the shear plane")] string timber2,
         [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber2,
         [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle2,
         [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var Timber1 = ExcelHelpers.GetTimberMaterialFromTag(timber1);
            var Timber2 = ExcelHelpers.GetTimberMaterialFromTag(timber2);

            return new TimberTimberDoubleShear(fastener, Timber1, ThicknessTimber1, angle1, Timber2, ThicknessTimber2, angle2, ropeEffect).Capacity;
        }

        [ExcelFunction(Description = "Return the characteristic shear capacity of a timber to timber connection in [N]. The value is for 1 shear plane",
            Name = "SDK.EC5.Connection.TimberTimberDoubleShearFailureMode",
            IsHidden = false,
            Category = "SDK.EC5.Connections_Utilities")]
        public static double TimberTimberDoubleShearFailureMode(
            [ExcelArgument(Description = "Tag representing the fastener (i.e Bolt_D10_Fu800)")] string FastenerTag,
            [ExcelArgument(Description = "First timber material used in the shear plane")] string timber1,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber1,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle1,
            [ExcelArgument(Description = "Second timber material used in the shear plane")] string timber2,
            [ExcelArgument(Description = "Timber thickness considered in the shear plane(t1 or t2 according to 8.2.3 (3))")] double ThicknessTimber2,
            [ExcelArgument(Description = "Load angle toward the timber grain in degree")] double angle2,
            [ExcelArgument(Description = "Boolean value which defines if the rope effect should be considered")] bool ropeEffect,
            [ExcelArgument(Description = "Failure mode to consider")] string failureMode)
        {
            var fastener = ExcelHelpers.GetFastenerFromTag(FastenerTag);
            var Timber1 = ExcelHelpers.GetTimberMaterialFromTag(timber1);
            var Timber2 = ExcelHelpers.GetTimberMaterialFromTag(timber2);

            var connection = new TimberTimberDoubleShear(fastener, Timber1, ThicknessTimber1, angle1, Timber2, ThicknessTimber2, angle2, ropeEffect);

            if (connection.FailureModes.Contains(failureMode)) return connection.Capacities[connection.FailureModes.IndexOf(failureMode)];
            else return -1.0;
        }


        #endregion



    }
}

