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
    public class MaterialTimberGlulam : IMaterial, IMaterialTimber
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

        public double B0 { get; set; }

        public double Bn { get; set; }
        #endregion

        #region constructor
        public MaterialTimberGlulam(string name)
        {

            if (Enum.GetNames(typeof(MaterialTimberGlulam.Grades)).Contains(name))
            {
                Grade = name;
            }
            else throw new ArgumentException(String.Format("The grade {0} is not present in the database, please look at the documentation", name));


            Fmyk = MaterialTimberGlulam.fmyk[Grade];
            Fmzk = MaterialTimberGlulam.fmzk[Grade];
            Ft0k = MaterialTimberGlulam.ft0k[Grade];
            Ft90k = MaterialTimberGlulam.ft90k[Grade];
            Fc0k = MaterialTimberGlulam.fc0k[Grade];
            Fc90k = MaterialTimberGlulam.fc90k[Grade];
            Fvk = MaterialTimberGlulam.fvk[Grade];
            Frk = MaterialTimberGlulam.frk[Grade];
            E0mean = MaterialTimberGlulam.e0mean[Grade];
            E90mean = MaterialTimberGlulam.e90mean[Grade];
            G0mean = MaterialTimberGlulam.gmean[Grade];
            E0_005 = MaterialTimberGlulam.e0_005[Grade];
            G0_005 = MaterialTimberGlulam.g0_005[Grade];
            RhoMean = MaterialTimberGlulam.rhoMean[Grade];
            RhoK = MaterialTimberGlulam.rhoK[Grade];
            B0 = MaterialTimberGlulam.b0;
            Bn = MaterialTimberGlulam.bn;

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }

        public MaterialTimberGlulam(Grades name)
        {
            Grade = name.ToString();


            Fmyk = MaterialTimberGlulam.fmyk[Grade];
            Fmzk = MaterialTimberGlulam.fmzk[Grade];
            Ft0k = MaterialTimberGlulam.ft0k[Grade];
            Ft90k = MaterialTimberGlulam.ft90k[Grade];
            Fc0k = MaterialTimberGlulam.fc0k[Grade];
            Fc90k = MaterialTimberGlulam.fc90k[Grade];
            Fvk = MaterialTimberGlulam.fvk[Grade];
            Frk = MaterialTimberGlulam.frk[Grade];
            E0mean = MaterialTimberGlulam.e0mean[Grade];
            E90mean = MaterialTimberGlulam.e90mean[Grade];
            G0mean = MaterialTimberGlulam.gmean[Grade];
            E0_005 = MaterialTimberGlulam.e0_005[Grade];
            G0_005 = MaterialTimberGlulam.g0_005[Grade];
            RhoMean = MaterialTimberGlulam.rhoMean[Grade];
            RhoK = MaterialTimberGlulam.rhoK[Grade];
            B0 = MaterialTimberGlulam.b0;
            Bn = MaterialTimberGlulam.bn;

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }
        #endregion

        #region Material properties

        public static TimberType type = TimberType.Glulam;

        //Material properties of Glulam according to DIN EN 14080:2013
        public enum Grades
        {
            GL20h,
            GL22h,
            GL24h,
            GL26h,
            GL28h,
            GL30h,
            GL32h,

            GL20c,
            GL22c,
            GL24c,
            GL26c,
            GL28c,
            GL30c,
            GL32c,
        };

        public static readonly Dictionary<string, double> fmyk = new Dictionary<string, double>

        {
            //Glulam
            {"GL20h", 20},
            {"GL22h", 22},
            {"GL24h", 24},
            {"GL26h", 26},
            {"GL28h", 28},
            {"GL30h", 30},
            {"GL32h", 32},

            {"GL20c", 20},
            {"GL22c", 22},
            {"GL24c", 24},
            {"GL26c", 26},
            {"GL28c", 28},
            {"GL30c", 30},
            {"GL32c", 32},

        };

        public static readonly Dictionary<string, double> fmzk = new Dictionary<string, double>

        {
             //Glulam
            {"GL20h", 20},
            {"GL22h", 22},
            {"GL24h", 24},
            {"GL26h", 26},
            {"GL28h", 28},
            {"GL30h", 30},
            {"GL32h", 32},

            {"GL20c", 20},
            {"GL22c", 22},
            {"GL24c", 24},
            {"GL26c", 26},
            {"GL28c", 28},
            {"GL30c", 30},
            {"GL32c", 32},

        };

        public static readonly Dictionary<string, double> ft0k = new Dictionary<string, double>
        {
             //Glulam
            {"GL20h", 16},
            {"GL22h", 17.6},
            {"GL24h", 19.2},
            {"GL26h", 20.8},
            {"GL28h", 22.3},
            {"GL30h", 24},
            {"GL32h", 25.6},

            {"GL20c", 15},
            {"GL22c", 16},
            {"GL24c", 17},
            {"GL26c", 19},
            {"GL28c", 19.5},
            {"GL30c", 19.5},
            {"GL32c", 19.5},
        };

        public static readonly Dictionary<string, double> ft90k = new Dictionary<string, double>
        {
             //Glulam
            {"GL20h", 0.5},
            {"GL22h", 0.5},
            {"GL24h", 0.5},
            {"GL26h", 0.5},
            {"GL28h", 0.5},
            {"GL30h", 0.5},
            {"GL32h", 0.5},

            {"GL20c", 0.5},
            {"GL22c", 0.5},
            {"GL24c", 0.5},
            {"GL26c", 0.5},
            {"GL28c", 0.5},
            {"GL30c", 0.5},
            {"GL32c", 0.5},

        };

        public static readonly Dictionary<string, double> fc0k = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 20},
            {"GL22h", 22},
            {"GL24h", 24},
            {"GL26h", 26},
            {"GL28h", 28},
            {"GL30h", 30},
            {"GL32h", 32},

            {"GL20c", 18.5},
            {"GL22c", 20},
            {"GL24c", 21.5},
            {"GL26c", 23.5},
            {"GL28c", 24},
            {"GL30c", 24.5},
            {"GL32c", 24.5},

        };

        public static readonly Dictionary<string, double> fc90k = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 2.5},
            {"GL22h", 2.5},
            {"GL24h", 2.5},
            {"GL26h", 2.5},
            {"GL28h", 2.5},
            {"GL30h", 2.5},
            {"GL32h", 2.5},

            {"GL20c", 2.5},
            {"GL22c", 2.5},
            {"GL24c", 2.5},
            {"GL26c", 2.5},
            {"GL28c", 2.5},
            {"GL30c", 2.5},
            {"GL32c", 2.5},


        };

        public static readonly Dictionary<string, double> fvk = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 3.5},
            {"GL22h", 3.5},
            {"GL24h", 3.5},
            {"GL26h", 3.5},
            {"GL28h", 3.5},
            {"GL30h", 3.5},
            {"GL32h", 3.5},

            {"GL20c", 3.5},
            {"GL22c", 3.5},
            {"GL24c", 3.5},
            {"GL26c", 3.5},
            {"GL28c", 3.5},
            {"GL30c", 3.5},
            {"GL32c", 3.5},


        };

        public static readonly Dictionary<string, double> frk = new Dictionary<string, double>
        {
             //Glulam
            {"GL20h", 1.2},
            {"GL22h", 1.2},
            {"GL24h", 1.2},
            {"GL26h", 1.2},
            {"GL28h", 1.2},
            {"GL30h", 1.2},
            {"GL32h", 1.2},

            {"GL20c", 1.2},
            {"GL22c", 1.2},
            {"GL24c", 1.2},
            {"GL26c", 1.2},
            {"GL28c", 1.2},
            {"GL30c", 1.2},
            {"GL32c", 1.2},


        };

        public static readonly Dictionary<string, double> e0mean = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 8400},
            {"GL22h", 10500},
            {"GL24h", 11500},
            {"GL26h", 12100},
            {"GL28h", 12600},
            {"GL30h", 13600},
            {"GL32h", 14200},

            {"GL20c", 10400},
            {"GL22c", 10400},
            {"GL24c", 11000},
            {"GL26c", 12000},
            {"GL28c", 12500},
            {"GL30c", 13000},
            {"GL32c", 13500},


        };

        public static readonly Dictionary<string, double> e90mean = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 300},
            {"GL22h", 300},
            {"GL24h", 300},
            {"GL26h", 300},
            {"GL28h", 300},
            {"GL30h", 300},
            {"GL32h", 300},

            {"GL20c", 300},
            {"GL22c", 300},
            {"GL24c", 300},
            {"GL26c", 300},
            {"GL28c", 300},
            {"GL30c", 300},
            {"GL32c", 300},


        };

        public static readonly Dictionary<string, double> gmean = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 650},
            {"GL22h", 650},
            {"GL24h", 650},
            {"GL26h", 650},
            {"GL28h", 650},
            {"GL30h", 650},
            {"GL32h", 650},

            {"GL20c", 650},
            {"GL22c", 650},
            {"GL24c", 650},
            {"GL26c", 650},
            {"GL28c", 650},
            {"GL30c", 650},
            {"GL32c", 650},

        };

        public static readonly Dictionary<string, double> e0_005 = new Dictionary<string, double>
        {
             //Glulam
            {"GL20h", 7000},
            {"GL22h", 8800},
            {"GL24h", 9600},
            {"GL26h", 10100},
            {"GL28h", 10500},
            {"GL30h", 11300},
            {"GL32h", 11800},

            {"GL20c", 8600},
            {"GL22c", 8600},
            {"GL24c", 9100},
            {"GL26c", 10000},
            {"GL28c", 10400},
            {"GL30c", 10800},
            {"GL32c", 11200},

        };

        public static readonly Dictionary<string, double> g0_005 = new Dictionary<string, double>
        {

            //Glulam
            {"GL20h", 540},
            {"GL22h", 540},
            {"GL24h", 540},
            {"GL26h", 540},
            {"GL28h", 540},
            {"GL30h", 540},
            {"GL32h", 540},

            {"GL20c", 540},
            {"GL22c", 540},
            {"GL24c", 540},
            {"GL26c", 540},
            {"GL28c", 540},
            {"GL30c", 540},
            {"GL32c", 540},

        };

        public static readonly Dictionary<string, double> rhoMean = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 370},
            {"GL22h", 410},
            {"GL24h", 420},
            {"GL26h", 445},
            {"GL28h", 460},
            {"GL30h", 480},
            {"GL32h", 490},

            {"GL20c", 390},
            {"GL22c", 390},
            {"GL24c", 400},
            {"GL26c", 420},
            {"GL28c", 420},
            {"GL30c", 430},
            {"GL32c", 440},


        };

        public static readonly Dictionary<string, double> rhoK = new Dictionary<string, double>
        {
            //Glulam
            {"GL20h", 340},
            {"GL22h", 370},
            {"GL24h", 385},
            {"GL26h", 405},
            {"GL28h", 425},
            {"GL30h", 430},
            {"GL32h", 440},

            {"GL20c", 355},
            {"GL22c", 355},
            {"GL24c", 365},
            {"GL26c", 385},
            {"GL28c", 390},
            {"GL30c", 390},
            {"GL32c", 400},


        };

        public static readonly double b0 = 0.65; //according to DIN EN 1995 - 1 - 2 table 3.1
        public static readonly double bn = 0.70; //according to DIN EN 1995 - 1 - 2 table 3.1
        #endregion
    }
}
