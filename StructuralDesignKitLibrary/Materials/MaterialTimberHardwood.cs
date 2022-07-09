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
    public class MaterialTimberHardwood : IMaterial, IMaterialTimber
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


        #region Constructors
        public MaterialTimberHardwood(string name)
        {

            if (Enum.GetNames(typeof(MaterialTimberHardwood.Grades)).Contains(name))
            {
                Grade = name;
            }
            else throw new ArgumentException(String.Format("The grade {0} is not present in the database, please look at the documentation", name));
            

            Fmyk = MaterialTimberHardwood.fmyk[Grade];
            Fmzk = MaterialTimberHardwood.fmzk[Grade];
            Ft0k = MaterialTimberHardwood.ft0k[Grade];
            Ft90k = MaterialTimberHardwood.ft90k[Grade];
            Fc0k = MaterialTimberHardwood.fc0k[Grade];
            Fc90k = MaterialTimberHardwood.fc90k[Grade];
            Fvk = MaterialTimberHardwood.fvk[Grade];
            Frk = MaterialTimberHardwood.frk[Grade];
            E0mean = MaterialTimberHardwood.e0mean[Grade];
            E90mean = MaterialTimberHardwood.e90mean[Grade];
            G0mean = MaterialTimberHardwood.gmean[Grade];
            E0_005 = MaterialTimberHardwood.e0_005[Grade];
            G0_005 = MaterialTimberHardwood.g0_005[Grade];
            RhoMean = MaterialTimberHardwood.rhoMean[Grade];
            RhoK = MaterialTimberHardwood.rhoK[Grade];

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }

        public MaterialTimberHardwood(Grades name)
        {
            Grade = name.ToString();


            Fmyk = MaterialTimberHardwood.fmyk[Grade];
            Fmzk = MaterialTimberHardwood.fmzk[Grade];
            Ft0k = MaterialTimberHardwood.ft0k[Grade];
            Ft90k = MaterialTimberHardwood.ft90k[Grade];
            Fc0k = MaterialTimberHardwood.fc0k[Grade];
            Fc90k = MaterialTimberHardwood.fc90k[Grade];
            Fvk = MaterialTimberHardwood.fvk[Grade];
            Frk = MaterialTimberHardwood.frk[Grade];
            E0mean = MaterialTimberHardwood.e0mean[Grade];
            E90mean = MaterialTimberHardwood.e90mean[Grade];
            G0mean = MaterialTimberHardwood.gmean[Grade];
            E0_005 = MaterialTimberHardwood.e0_005[Grade];
            G0_005 = MaterialTimberHardwood.g0_005[Grade];
            RhoMean = MaterialTimberHardwood.rhoMean[Grade];
            RhoK = MaterialTimberHardwood.rhoK[Grade];

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }
        #endregion


        #region Material properties

        public static TimberType type = TimberType.Hardwood;

        //Material properties of Hardwood according to DIN EN 338:2016
        public enum Grades
        {
            D30,
            D35,
            D40,
            D60,
        };

        public static readonly Dictionary<string, double> fmyk = new Dictionary<string, double>

        {
            //Hardwood
            {"D30", 30},
            {"D35", 35},
            {"D40", 40},
            {"D60", 60},

        };

        public static readonly Dictionary<string, double> fmzk = new Dictionary<string, double>

        {
             //Hardwood
            {"D30", 30},
            {"D35", 35},
            {"D40", 40},
            {"D60", 60},
        };

        public static readonly Dictionary<string, double> ft0k = new Dictionary<string, double>
        {
             //Hardwood
            {"D30", 18},
            {"D35", 21},
            {"D40", 24},
            {"D60", 36},
        };

        public static readonly Dictionary<string, double> ft90k = new Dictionary<string, double>
        {
             //Hardwood
            {"D30", 0.6},
            {"D35", 0.6},
            {"D40", 0.6},
            {"D60", 0.6},

        };

        public static readonly Dictionary<string, double> fc0k = new Dictionary<string, double>
        {
             //Hardwood
            {"D30", 24},
            {"D35", 25},
            {"D40", 27},
            {"D60", 33},

        };

        public static readonly Dictionary<string, double> fc90k = new Dictionary<string, double>
        {
             //Hardwood
            {"D30", 5.3},
            {"D35", 5.4},
            {"D40", 5.5},
            {"D60", 10.5},


        };

        public static readonly Dictionary<string, double> fvk = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 3.9},
            {"D35", 4.1},
            {"D40", 4.2},
            {"D60", 4.8},


        };

        public static readonly Dictionary<string, double> frk = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 0},
            {"D35", 0},
            {"D40", 0},
            {"D60", 0},


        };

        public static readonly Dictionary<string, double> e0mean = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 11000},
            {"D35", 12000},
            {"D40", 13000},
            {"D60", 17000},


        };

        public static readonly Dictionary<string, double> e90mean = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 730},
            {"D35", 800},
            {"D40", 870},
            {"D60", 1130},


        };

        public static readonly Dictionary<string, double> gmean = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 690},
            {"D35", 750},
            {"D40", 810},
            {"D60", 1060},

        };

        public static readonly Dictionary<string, double> e0_005 = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 9200},
            {"D35", 10100},
            {"D40", 10900},
            {"D60", 14300},

        };

        public static readonly Dictionary<string, double> g0_005 = new Dictionary<string, double>
        {
            //According to DIN EN 1995-1-1/NA §NCI Zu 3.2 "Vollholz" (NA.7)  : G05 = 2/3 * Gmean

            //Hardwood
            {"D30", 690 * 2/3},
            {"D35", 750 * 2/3},
            {"D40", 810 * 2/3},
            {"D60", 1060 * 2/3},

        };

        public static readonly Dictionary<string, double> rhoMean = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 640},
            {"D35", 650},
            {"D40", 660},
            {"D60", 840},


        };

        public static readonly Dictionary<string, double> rhoK = new Dictionary<string, double>
        {
            //Hardwood
            {"D30", 530},
            {"D35", 540},
            {"D40", 550},
            {"D60", 700},


        };
        #endregion
    }
}
