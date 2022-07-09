using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StructuralDesignKitLibrary.EC5;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;

namespace StructuralDesignKitLibrary.Materials
{
    public class MaterialTimberBaubuche : IMaterial, IMaterialTimber
    {

        #region Object properties

        public string Grade { get; set; }

        public double Density { get; set; }

        public double E { get; set; }

        public double G { get; set; }

        public TimberType Type { get { return type; } }

        public double Fmyk { get; set; }

        public double Fmzk { get; set; }

        public double Ft0k { get; set; }

        public double Ft90k { get; set; }

        public double Fc0k { get; set; }

        public double Fc90k { get; set; }

        public double Fvk { get; set; }

        public double Frk { get; set; }

        public double E0mean { get; set; }

        public double E90mean { get; set; }

        public double G0mean { get; set; }

        public double E0_005 { get; set; }

        public double G0_005 { get; set; }

        public double RhoMean { get; set; }

        public double RhoK { get; set; }
        #endregion

        #region constructor
        public MaterialTimberBaubuche(string name)
        {

            if (Enum.GetNames(typeof(MaterialTimberBaubuche.Grade_Baubuche)).Contains(name))
            {
                Grade = name;
            }
            else throw new ArgumentException(String.Format("The grade {0} is not present in the database, please look at the documentation", name));


            Fmyk = MaterialTimberBaubuche.fmyk[Grade];
            Fmzk = MaterialTimberBaubuche.fmzk[Grade];
            Ft0k = MaterialTimberBaubuche.ft0k[Grade];
            Ft90k = MaterialTimberBaubuche.ft90k[Grade];
            Fc0k = MaterialTimberBaubuche.fc0k[Grade];
            Fc90k = MaterialTimberBaubuche.fmzk[Grade];
            Fvk = MaterialTimberBaubuche.fmzk[Grade];
            Frk = MaterialTimberBaubuche.fmzk[Grade];
            E0mean = MaterialTimberBaubuche.fmzk[Grade];
            E90mean = MaterialTimberBaubuche.fmzk[Grade];
            G0mean = MaterialTimberBaubuche.fmzk[Grade];
            E0_005 = MaterialTimberBaubuche.fmzk[Grade];
            G0_005 = MaterialTimberBaubuche.fmzk[Grade];
            RhoMean = MaterialTimberBaubuche.fmzk[Grade];
            RhoK = MaterialTimberBaubuche.fmzk[Grade];

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }

        public MaterialTimberBaubuche(Grade_Baubuche name)
        {
            Grade = name.ToString();


            Fmyk = MaterialTimberBaubuche.fmyk[Grade];
            Fmzk = MaterialTimberBaubuche.fmzk[Grade];
            Ft0k = MaterialTimberBaubuche.ft0k[Grade];
            Ft90k = MaterialTimberBaubuche.ft90k[Grade];
            Fc0k = MaterialTimberBaubuche.fc0k[Grade];
            Fc90k = MaterialTimberBaubuche.fmzk[Grade];
            Fvk = MaterialTimberBaubuche.fmzk[Grade];
            Frk = MaterialTimberBaubuche.fmzk[Grade];
            E0mean = MaterialTimberBaubuche.fmzk[Grade];
            E90mean = MaterialTimberBaubuche.fmzk[Grade];
            G0mean = MaterialTimberBaubuche.fmzk[Grade];
            E0_005 = MaterialTimberBaubuche.fmzk[Grade];
            G0_005 = MaterialTimberBaubuche.fmzk[Grade];
            RhoMean = MaterialTimberBaubuche.fmzk[Grade];
            RhoK = MaterialTimberBaubuche.fmzk[Grade];

            //according to design guide, density for load calculation 850Kg/m³
            Density = 850;
            E = E0mean;
            G = G0mean;

        }
        #endregion

        #region Material properties

        public static TimberType type = TimberType.Baubuche;

        //Material properties of Baubuche GL75h according to ETA-14/0354 of 11.07.2018
        public enum Grade_Baubuche
        {
            GL75h_Cl1,
            GL75h_Cl2,
        };

        public static readonly Dictionary<string, double> fmyk = new Dictionary<string, double>

        {
            //Baubuche
            {"GL75h_Cl1", 75},
            {"GL75h_Cl2", 75},

        };

        public static readonly Dictionary<string, double> fmzk = new Dictionary<string, double>

        {
            //Baubuche
            {"GL75h_Cl1", 75},
            {"GL75h_Cl2", 75},

        };

        public static readonly Dictionary<string, double> ft0k = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 60},
            {"GL75h_Cl2", 60},
        };

        public static readonly Dictionary<string, double> ft90k = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 0.6},
            {"GL75h_Cl2", 0.6},

        };

        public static readonly Dictionary<string, double> fc0k = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 59.4},
            {"GL75h_Cl2", 49.5},

        };

        public static readonly Dictionary<string, double> fc90k = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 14.8},
            {"GL75h_Cl2", 12.3},


        };

        public static readonly Dictionary<string, double> fvk = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 4.5},
            {"GL75h_Cl2", 4.5},


        };

        public static readonly Dictionary<string, double> frk = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 0},
            {"GL75h_Cl2", 0},


        };

        public static readonly Dictionary<string, double> e0mean = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 16800},
            {"GL75h_Cl2", 16800},


        };

        public static readonly Dictionary<string, double> e90mean = new Dictionary<string, double>
        {
         //Baubuche
            {"GL75h_Cl1", 470},
            {"GL75h_Cl2", 470},


        };

        public static readonly Dictionary<string, double> gmean = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 850},
            {"GL75h_Cl2", 850},

        };

        public static readonly Dictionary<string, double> e0_005 = new Dictionary<string, double>
        {

            //Baubuche
            {"GL75h_Cl1", 15300},
            {"GL75h_Cl2", 15300},

        };

        public static readonly Dictionary<string, double> g0_005 = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 760},
            {"GL75h_Cl2", 760},

        };

        public static readonly Dictionary<string, double> rhoMean = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 800},
            {"GL75h_Cl2", 800},


        };

        public static readonly Dictionary<string, double> rhoK = new Dictionary<string, double>
        {
            //Baubuche
            {"GL75h_Cl1", 730},
            {"GL75h_Cl2", 730},


        };
        #endregion


        /// <summary>
        /// Update the material properties based on the size modification factors in the 
        /// ETA-14/0354 of 20.09.2021 and in Manual for design and structural calculation in accordance with Eurocode 5 - 3rd revised edition
        /// from Hans Joachim Blass, Johannes Streib
        /// </summary>
        /// <param name="b">beam width</param>
        /// <param name="h">beam height - represents the Z axis of the beam, where lamellas are stacked</param>
        [Description("Update the material properties based on the size modification factor")]
        public void UpdateBaubucheProperties(int b, int h)
        {
            //Update bending strength flatwise (Y axis):
            Fmyk *= Math.Pow(600 / h, 0.1);

            //Update bending strength edgewise (Z axis) according to design guide P.11:
            if (b > 300)Fmzk *= Math.Pow(300 / h, 0.12);
            if (b > 1200) throw new Exception("Baubuche Block gluing is limited to 1200mm according to ETA-14/0354");

            //Update tension strength:
            Ft0k*= Math.Pow(600 / Math.Max(h, b), 0.1);

            //Update compression parallel to the grain strength:
            var kc = Math.Min(1.18, 0.0009 * h + 0.892);
            if (h > 120) Fc0k *= kc;

            //Update shear strength:
            var kvh = Math.Pow(600 / h, 0.13);
            Fvk *= kvh;
        }
    }
}
