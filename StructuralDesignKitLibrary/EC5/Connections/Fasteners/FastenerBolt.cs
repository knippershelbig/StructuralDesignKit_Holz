using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using StructuralDesignKitLibrary.Connections.TimberTimberShear;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.EC5.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
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
        public double a2min { get; set;}
        public double a3tmin { get;set; }
        public double a3cmin { get;set; }
        public double a4tmin { get;set; }
        public double a4cmin { get;set; }
        public double K90 { get; set; }

        #endregion


        #region constructor
        public FastenerBolt(double diameter, double fuk)
        {
            Type = EC5_Utilities.FastenerType.Dowel;
            if (diameter <= 30) Diameter = diameter;
            else throw new Exception("According to EN 1995-1-1 §8.5(2), the maximum bolt diameter allowed is 30mm");
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
            if (angle <= 150 && angle < 210) return 4*Diameter;
            else return (1+6*Math.Sin(AngleRad))*Diameter;
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
            double AngleRad = angle * Math.PI / 180;
            double fh0k = 0.082 * (1 - 0.01 * Diameter) * timber.RhoK; //EN 1995-1-1 Eq (8.32)
            K90 = ComputeK90(timber);
            Fhk = fh0k / (K90 * Math.Pow(Math.Sin(AngleRad), 2) + Math.Pow(Math.Cos(AngleRad), 2)); //EN 1995-1-1 Eq (8.32)

        }
        /// <summary>
        /// Bolt withdrawal capacity according to EN 1995-1-1 §8.5.2
        /// </summary>
        /// <param name="ConnectionType"></param>
        /// <exception cref="Exception"></exception>
        public void ComputeWithdrawalStrength(IShearCapacity ConnectionType)
        {
            if (ConnectionType is SteelSingleInnerPlate )
            {
                var connection = ConnectionType as SteelSingleInnerPlate;

                //Washer diameter considered as 3x the bolt diameter
                double boltTensileCapacity = Math.PI * Math.Pow(Diameter, 2) / 4 * Fuk;
                double washerCompressiveCapacity = Math.PI/4 * (Math.Pow(Diameter * 3, 2)- Math.Pow(Diameter +2, 2)) * connection.Timber.Fc90k * 3;
                WithdrawalStrength = Math.Min(boltTensileCapacity, washerCompressiveCapacity);
            }
            else if (ConnectionType is TimberTimberSingleShear)
            {
                var connection = ConnectionType as TimberTimberSingleShear;
                double fc90k = Math.Min(connection.Timber1.Fc90k, connection.Timber2.Fc90k);

                //Washer diameter considered as 3x the bolt diameter
                double boltTensileCapacity = Math.PI * Math.Pow(Diameter, 2) / 4 * Fuk;
                double washerCompressiveCapacity = Math.PI / 4 * (Math.Pow(Diameter * 3, 2) - Math.Pow(Diameter + 2, 2)) * fc90k * 3;
                WithdrawalStrength = Math.Min(boltTensileCapacity, washerCompressiveCapacity);
            }
            else throw new Exception("Bolt Withdrawal capacity not yet implemented");
        }


        public double ComputeEffectiveFastener(int n, double a1)
        {
            throw new NotImplementedException();
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
