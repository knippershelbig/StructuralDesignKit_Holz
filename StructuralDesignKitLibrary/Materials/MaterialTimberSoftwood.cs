using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using StructuralDesignKitLibrary.EC5;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;

namespace StructuralDesignKitLibrary.Materials
{
    [Description("Structural timber classs for Softwood according to EN 1995 and EN 338:2016")]
    public class MaterialTimberSoftwood : IMaterial, IMaterialTimber
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

        public MaterialTimberSoftwood(string name)
        {
            
            if (Enum.GetNames(typeof(MaterialTimberSoftwood.Grades)).Contains(name))
            {
                Grade = name;
            }
            else throw new ArgumentException(String.Format("The grade {0} is not present in the database, please look at the documentation",name));


            Fmyk = MaterialTimberSoftwood.fmyk[Grade];
            Fmzk = MaterialTimberSoftwood.fmzk[Grade];
            Ft0k = MaterialTimberSoftwood.ft0k[Grade];
            Ft90k = MaterialTimberSoftwood.ft90k[Grade];
            Fc0k = MaterialTimberSoftwood.fc0k[Grade];
            Fc90k = MaterialTimberSoftwood.fc90k[Grade];
            Fvk = MaterialTimberSoftwood.fvk[Grade];
            Frk = MaterialTimberSoftwood.frk[Grade];
            E0mean = MaterialTimberSoftwood.e0mean[Grade];
            E90mean = MaterialTimberSoftwood.e90mean[Grade];
            G0mean = MaterialTimberSoftwood.gmean[Grade];
            E0_005 = MaterialTimberSoftwood.e0_005[Grade];
            G0_005 = MaterialTimberSoftwood.g0_005[Grade];
            RhoMean = MaterialTimberSoftwood.rhoMean[Grade];
            RhoK = MaterialTimberSoftwood.rhoK[Grade];

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }

        public MaterialTimberSoftwood(Grades name)
        {
            Grade = name.ToString();


            Fmyk = MaterialTimberSoftwood.fmyk[Grade];
            Fmzk = MaterialTimberSoftwood.fmzk[Grade];
            Ft0k = MaterialTimberSoftwood.ft0k[Grade];
            Ft90k = MaterialTimberSoftwood.ft90k[Grade];
            Fc0k = MaterialTimberSoftwood.fc0k[Grade];
            Fc90k = MaterialTimberSoftwood.fc90k[Grade];
            Fvk = MaterialTimberSoftwood.fvk[Grade];
            Frk = MaterialTimberSoftwood.frk[Grade];
            E0mean = MaterialTimberSoftwood.e0mean[Grade];
            E90mean = MaterialTimberSoftwood.e90mean[Grade];
            G0mean = MaterialTimberSoftwood.gmean[Grade];
            E0_005 = MaterialTimberSoftwood.e0_005[Grade];
            G0_005 = MaterialTimberSoftwood.g0_005[Grade];
            RhoMean = MaterialTimberSoftwood.rhoMean[Grade];
            RhoK = MaterialTimberSoftwood.rhoK[Grade];

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }
        #endregion

        #region Material properties

        public static TimberType type = TimberType.Softwood;

        //Material properties of Softwood according to DIN EN 338:2016
        public enum Grades
        {
            C16,
            C24,
            C30,
            C35,
            C40,
        };

        public static readonly Dictionary<string, double> fmyk = new Dictionary<string, double>

        {
            //Softwood
            {"C16", 16},
            {"C24", 24},
            {"C30", 30},
            {"C35", 35},
            {"C40", 40}

        };

        public static readonly Dictionary<string, double> fmzk = new Dictionary<string, double>

        {
            //Softwood
            {"C16", 16},
            {"C24", 24},
            {"C30", 30},
            {"C35", 35},
            {"C40", 40}

        };

        public static readonly Dictionary<string, double> ft0k = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 8.5},
            {"C24", 14.5},
            {"C30", 19},
            {"C35", 22.5},
            {"C40", 26},
        };

        public static readonly Dictionary<string, double> ft90k = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 0.4},
            {"C24", 0.4},
            {"C30", 0.4},
            {"C35", 0.4},
            {"C40", 0.4},

        };

        public static readonly Dictionary<string, double> fc0k = new Dictionary<string, double>
        {
           //Softwood
            {"C16", 17},
            {"C24", 21},
            {"C30", 24},
            {"C35", 25},
            {"C40", 27},

        };

        public static readonly Dictionary<string, double> fc90k = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 2.2},
            {"C24", 2.5},
            {"C30", 2.7},
            {"C35", 2.7},
            {"C40", 2.8},


        };

        public static readonly Dictionary<string, double> fvk = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 3.2},
            {"C24", 4},
            {"C30", 4},
            {"C35", 4},
            {"C40", 4},


        };

        public static readonly Dictionary<string, double> frk = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 0},
            {"C24", 0},
            {"C30", 0},
            {"C35", 0},
            {"C40", 0},


        };

        public static readonly Dictionary<string, double> e0mean = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 8000},
            {"C24", 11000},
            {"C30", 12000},
            {"C35", 13000},
            {"C40", 14000},


        };

        public static readonly Dictionary<string, double> e90mean = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 270},
            {"C24", 370},
            {"C30", 400},
            {"C35", 430},
            {"C40", 470},


        };

        public static readonly Dictionary<string, double> gmean = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 500},
            {"C24", 690},
            {"C30", 750},
            {"C35", 810},
            {"C40", 880},

        };

        public static readonly Dictionary<string, double> e0_005 = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 5400},
            {"C24", 7400},
            {"C30", 8000},
            {"C35", 8700},
            {"C40", 9400},

        };

        public static readonly Dictionary<string, double> g0_005 = new Dictionary<string, double>
        {
            //According to DIN EN 1995-1-1/NA §NCI Zu 3.2 "Vollholz" (NA.7)  : G05 = 2/3 * Gmean

            //Softwood
            {"C16", 500 * 2/3},
            {"C24", 690 * 2/3},
            {"C30", 750 * 2/3},
            {"C35", 810 * 2/3},
            {"C40", 880 * 2/3},

        };

        public static readonly Dictionary<string, double> rhoMean = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 370},
            {"C24", 420},
            {"C30", 460},
            {"C35", 470},
            {"C40", 480},


        };

        public static readonly Dictionary<string, double> rhoK = new Dictionary<string, double>
        {
            //Softwood
            {"C16", 310},
            {"C24", 350},
            {"C30", 380},
            {"C35", 390},
            {"C40", 400},


        };
        #endregion
    }
}
