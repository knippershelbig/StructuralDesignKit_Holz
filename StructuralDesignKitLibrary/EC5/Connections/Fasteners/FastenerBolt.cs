using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using StructuralDesignKitLibrary.Connections.TimberTimberShear;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.EC5.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.Utilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.Fasteners
{
    /// <summary>
    /// Dowel fastener according to EN 1995-1-1 §8.6
    /// </summary>
    public class FastenerBolt : IFastener
    {

        #region Properties
        public EC5_Utilities.FastenerType Type { get; }
        public double Diameter { get; set; }
        public double Fuk { get; set; }
        public double MyRk { get; set; }
        public double Fhk { get; set; }
        public double WithdrawalStrength { get; set; }
        public double MaxJohansenPart { get; set; }
        public double a1min { get; set; }
        public double a2min { get; set; }
        public double a3tmin { get; set; }
        public double a3cmin { get; set; }
        public double a4tmin { get; set; }
        public double a4cmin { get; set; }
        public double K90 { get; set; }

        #endregion


        #region constructor
        public FastenerBolt(double diameter, double fuk)
        {
            Type = EC5_Utilities.FastenerType.Dowel;
            if (diameter <= 30 && diameter >= 6) Diameter = diameter;
            else throw new Exception("According to EN 1995-1-1 §8.5(2), the bolt diameter should be between 6mm and 30mm");
            Fuk = fuk;
            MyRk = 0.3 * Fuk * Math.Pow(Diameter, 2.6); //EN 1995-1-1 Eq (8.30)
            MaxJohansenPart = 0.25;
        }
        #endregion

        //Spacings according to EN 1995-1-1 Table 8.4
        #region define spacings

        /// <summary>
        /// Define the minimum spacing to alongside the grain in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the minimum spacing to alongside the grain in mm")]
        private double DefineA1Min(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            return (4 + Math.Abs(Math.Cos(AngleRad))) * Diameter;
        }

        /// <summary>
        /// Define the minimum spacing perpendicular to grain in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the minimum spacing perpendicular to grain in mm")]
        private double DefineA2Min(double angle)
        {
            return 4 * Diameter;
        }


        /// <summary>
        /// Define the Minimum spacing to loaded end in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the Minimum spacing to loaded end in mm")]
        private double DefineA3tMin(double angle)
        {
            return Math.Max(7 * Diameter, 80);
        }

        /// <summary>
        /// Define the Minimum spacing to unloaded end in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the Minimum spacing to unloaded end in mm")]
        private double DefineA3cMin(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            if (angle <= 150 && angle < 210) return 4 * Diameter;
            else return (1 + 6 * Math.Sin(AngleRad)) * Diameter;
        }

        /// <summary>
        /// Define the Minimum spacing to loaded edge in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the Minimum spacing to loaded edge in mm")]
        private double DefineA4tMin(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            return Math.Max((2 + 2 * Math.Sin(AngleRad)) * Diameter, 3 * Diameter);

        }

        /// <summary>
        /// Define the minimum spacing to unloaded edge in mm
        /// </summary>
        /// <returns></returns>
        [Description("Define the minimum spacing to unloaded edge in mm")]
        private double DefineA4cMin()
        {
            return 3 * Diameter;
        }

        #endregion


        /// <summary>
        /// Computes the K90 factor according to EN 1995-1-1 Eq (8.33) - K90 is a factor taking into consideration the splitting risk and the degree of compressive deformation
        /// </summary>
        /// <param name="timber"></param>
        /// <returns></returns>
        [Description("Computes the K90 factor according to EN 1995-1-1 Eq (8.33) - K90 is a factor taking into consideration the splitting risk and the degree of compressive deformation")]
        public double ComputeK90(IMaterialTimber timber)
        {
            double k90 = 0;

            switch (timber.Type)
            {
                case EC5_Utilities.TimberType.Softwood:
                    k90 = 1.35 + 0.015 * Diameter;
                    break;
                case EC5_Utilities.TimberType.Hardwood:
                    k90 = 0.9 + 0.015 * Diameter;
                    break;
                case EC5_Utilities.TimberType.Glulam:
                    k90 = 1.35 + 0.015 * Diameter;
                    break;
                case EC5_Utilities.TimberType.LVL:
                    k90 = 1.3 + 0.015 * Diameter;
                    break;
                case EC5_Utilities.TimberType.Baubuche:
                    k90 = 0.9 + 0.015 * Diameter;
                    break;
            }
            return k90;
        }


        public void ComputeEmbedmentStrength(IMaterialTimber timber, double angle)
        {
            List<string> CoveredTimber = new List<string>() { "Softwood", "Hardwood", "Glulam", "LVL", "Baubuche" };
            if (CoveredTimber.Contains(timber.Type.ToString()))
            {
                double AngleRad = angle * Math.PI / 180;
                double fh0k = 0.082 * (1 - 0.01 * Diameter) * timber.RhoK; //EN 1995-1-1 Eq (8.32)
                K90 = ComputeK90(timber);
                Fhk = fh0k / (K90 * Math.Pow(Math.Sin(AngleRad), 2) + Math.Pow(Math.Cos(AngleRad), 2)); //EN 1995-1-1 Eq (8.32)
            }
            else throw new Exception("Timber type not yet covered in the SDK");

        }
        /// <summary>
        /// Bolt withdrawal capacity according to EN 1995-1-1 §8.5.2
        /// </summary>
        /// <param name="ConnectionType"></param>
        /// <exception cref="Exception"></exception>
        public void ComputeWithdrawalStrength(IShearCapacity ConnectionType)
        {

            if (ConnectionType is SingleInnerSteelPlate)
            {
                var connection = ConnectionType as SingleInnerSteelPlate;
                double fc90k = connection.Timber.Fc90k;
                WithdrawalStrength = Math.Min(ComputeBoltTensileStrength(), ComputeWasherCompressiveStrength(fc90k));
            }

            else if (ConnectionType is TimberTimberSingleShear)
            {
                var connection = ConnectionType as TimberTimberSingleShear;
                double fc90k = Math.Min(connection.Timber1.Fc90k, connection.Timber2.Fc90k);
                WithdrawalStrength = Math.Min(ComputeBoltTensileStrength(), ComputeWasherCompressiveStrength(fc90k));
            }

            else if (ConnectionType is SingleOuterSteelPlate)
            {
                var connection = ConnectionType as SingleOuterSteelPlate;
                double fc90k = connection.Timber.Fc90k;
                List<double> withdrawalStrengths = new List<double>();
                withdrawalStrengths.Add(ComputeBoltTensileStrength());
                withdrawalStrengths.Add(ComputeWasherCompressiveStrength(fc90k));
                withdrawalStrengths.Add(ComputeWasherCompressiveStrength(fc90k,connection.SteelPlateThickness));

                WithdrawalStrength = withdrawalStrengths.Min();
            }

            else throw new Exception("Bolt Withdrawal capacity not yet implemented");
        }

        private double ComputeWasherCompressiveStrength(double fc90k, double plateThickness = -1)
        {
            //If a steel plate is considered as washer, contact area according to EC5 §8.5.2 (3)
            if (plateThickness > 0)
            {
                double area1 = Math.PI / 4 * (Math.Pow(12 * plateThickness, 2) - Math.Pow(Diameter + 2, 2));
                double area2 = Math.PI / 4 * (Math.Pow(Diameter * 4, 2) - Math.Pow(Diameter + 2, 2));
                return Math.Min(area1, area2) * 3 * fc90k;
            }

            //Washer diameter considered as 3x the bolt diameter and drilling diameter diam+2;
            else return Math.PI / 4 * (Math.Pow(Diameter * 3, 2) - Math.Pow(Diameter + 2, 2)) * fc90k * 3;
        }

        /// <summary>
        /// Compute the tensile strength of a bolt according to EC3-1.8 §3.6.1 table 3.4
        /// </summary>
        /// <returns></returns>
        private double ComputeBoltTensileStrength()
        {
            //To replace by a simple dictionary...
            //2nd Order polynomial equation to go from diameter to Stress area (threaded part) As[mm2] based on values from diameter 6 to 30mm
            //Provides results with a margin error of 3% 
            double aeff = (-0.00007 * Math.Pow(Diameter, 2) + 0.00575 * Diameter + 0.69089) * Math.PI * Math.Pow(Diameter, 2) / 4;

            return 0.9 * aeff * Fuk / 1.25; //consider k2 = 0.9 and Ym2 = 1.25
        }


        public double ComputeEffectiveNumberOfFastener(int n, double a1, double angle)
        {
            if (n == 1) return 1;
            else
            {
                double nef_0 = Math.Min(n, Math.Pow(n, 0.9) * Math.Pow(a1 / (13 * Diameter), 0.25));
                return SDKUtilities.LinearInterpolation(angle, 0, nef_0, 90, n);
            }
        }

        public void ComputeSpacings(double angle)
        {
            a1min = DefineA1Min(angle);
            a2min = DefineA2Min(angle);
            a3tmin = DefineA3tMin(angle);
            a3cmin = DefineA3cMin(angle);
            a4tmin = DefineA4tMin(angle);
            a4cmin = DefineA4cMin();
        }

    }
}
