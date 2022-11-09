using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.EC5;
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
    public class FasternerDowel : IFastener
    {

        #region Properties
        public EC5_Utilities.FasternType Type { get; }
        public double Diameter { get; set; }
        public double Fuk { get; set; }
        public double MyRk { get; set; }
        public double FhAlphaK { get; set; }
        public double WithdrawalStrength { get; set; }
        public double MaxJohansenPart { get; set; }
        public double a1min { get; }
        public double a2min { get; }
        public double a3tmin { get; }
        public double a3cmin { get; }
        public double a4tmin { get; }
        public double a4cmin { get; }

        public double K90 { get; set; }

        #endregion


        #region constructor
        public FasternerDowel(double diameter, double fuk)
        {
            Type = EC5_Utilities.FasternType.Dowel;
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

        public double DefineA1Min(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            return (3 + 2 * Math.Abs(Math.Cos(AngleRad))) * Diameter;
        }

        public double DefineA2Min(double angle)
        {
            return 3 * Diameter;
        }

        public double DefineA3tMin(double angle)
        {
            return Math.Max(7 * Diameter, 80);
        }

        public double DefineA3cMin(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            if (angle <= 150 && angle < 210) return Math.Max(3.5 * Diameter, 40);
            else return a3tmin;
        }

        public double DefineA4tMin(double angle)
        {
            double AngleRad = angle * Math.PI / 180;
            return Math.Max((2 + 2 * AngleRad) * Diameter, 3 * Diameter);

        }

        public double DefineA4cMin()
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
            FhAlphaK = fh0k / (K90 * Math.Pow(Math.Sin(AngleRad), 2) + Math.Pow(Math.Cos(AngleRad), 2)); //EN 1995-1-1 Eq (8.32)

        }



        public double ComputeEffectiveFastener(int n, double a1)
        {
            throw new NotImplementedException();
        }

        public void ComputeWithdrawalStrength(ISteelTimberShear ConnectionType)
        {
            WithdrawalStrength = 0;
        }
    }
}
