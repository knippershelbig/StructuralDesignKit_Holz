using StructuralDesignKitLibrary.Connections.Interface;
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
    public class FastenerDowel : IFastener
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
        public FastenerDowel(double diameter, double fuk)
        {
            Type = EC5_Utilities.FastenerType.Dowel;
            if (diameter > 6 && diameter < 30) Diameter = diameter;
            else throw new Exception("According to EN 1995-1-1 §8.6(2), the dowel diameter should be greater than 6mm and smaller than 30mm");
            Fuk = fuk;
            MyRk = 0.3 * Fuk * Math.Pow(Diameter, 2.6); //EN 1995-1-1 Eq (8.30)
            MaxJohansenPart = 0;
            WithdrawalStrength = 0;
        }
        #endregion

        //Spacings according to EN 1995-1-1 Table 8.5
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
            return (3 + 2 * Math.Abs(Math.Cos(AngleRad))) * Diameter;
        }

        /// <summary>
        /// Define the minimum spacing perpendicular to grain in mm
        /// </summary>
        /// <param name="angle">angle to grain in Degree</param>
        /// <returns></returns>
        [Description("Define the minimum spacing perpendicular to grain in mm")]
        private double DefineA2Min(double angle)
        {
            return 3 * Diameter;
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
            if (angle <= 150 && angle < 210) return Math.Max(3.5 * Diameter, 40);
            else return a3tmin;
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

        public double ComputeEffectiveNumberOfFastener(int n, double a1, double angle)
        {
            if (n == 1) return 1;
            else
            {
                double nef_0 = Math.Min(n, Math.Pow(n, 0.9) * Math.Pow(a1 / (13 * Diameter), 0.25));

                int angleFirstQuadrant = SDKUtilities.ComputeAngleToFirstQuadrant(angle);

                return SDKUtilities.LinearInterpolation(Convert.ToDouble(angleFirstQuadrant), 0, nef_0, 90, n);
            }
        }

        public void ComputeWithdrawalStrength(IShearCapacity ConnectionType)
        {
            WithdrawalStrength = 0;
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
