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
    [Description("Material for CLT definition based on ETAs")]
    public class MaterialCLT : IMaterial, IMaterialCLT
    {

        #region Object properties

        public string Grade { get; set; }
        public TimberType Type { get; set; }
        public double Density { get; set; }
        public double E { get; set; }
        public double G { get; set; }
        public double Fmk { get; set; }
        public double Ft0k { get; set; }
        public double Ft90k { get; set; }
        public double Fc0k { get; set; }
        public double Fc90k { get; set; }
        public double Fvk { get; set; }
        public double Frk { get; set; }
        public double E0mean { get; set; }
        public double E90mean { get; set; }
        public double G0meanPerp { get; set; }
        public double G0meanInPlane { get; set; }
        public double G90mean { get; set; }
        public double E0_005 { get; set; }
        public double E90_005 { get; set; }
        public double G0_005Perp { get; set; }
        public double G0_005InPlane { get; set; }
        public double G90_005 { get; set; }
        public double RhoMean { get; set; }
        public double RhoK { get; set; }
        public double B1_Vertical { get; set; }
        public double B1_Vertical_Increased { get; set; }
        public double B2_Vertical { get; set; }
        public double B2_Vertical_Increased { get; set; }
        public double B1_Horizontal { get; set; }
        public double B1_Horizontal_Increased { get; set; }
        public double B2_Horizontal { get; set; }
        public double B2_Horizontal_Increased { get; set; }

        #endregion

        #region constructor
        public MaterialCLT(ProducerGrade name)
        {
            DefineProperties(name);
        }

        public MaterialCLT(string producer_grade)
        {

            if (Enum.GetNames(typeof(MaterialCLT.ProducerGrade)).Contains(producer_grade))
            {
                Grade = producer_grade;
                ProducerGrade grade;
                ProducerGrade.TryParse(Grade, out grade);
                DefineProperties(grade);
            }
            else throw new ArgumentException(String.Format("The grade {0} is not present in the database, please look at the documentation", producer_grade));
        }

        private void DefineProperties(ProducerGrade name)
        {
            Grade = name.ToString();
            Density = MaterialCLT.rhoMean[Grade];
            Fmk = MaterialCLT.fmk[Grade];
            Ft0k = MaterialCLT.ft0k[Grade];
            Ft90k = MaterialCLT.ft90k[Grade];
            Fc0k = MaterialCLT.fc0k[Grade];
            Fc90k = MaterialCLT.fc90k[Grade];
            Fvk = MaterialCLT.fvk[Grade];
            Frk = MaterialCLT.frk[Grade];
            E0mean = MaterialCLT.e0mean[Grade];
            E90mean = MaterialCLT.e90mean[Grade];
            G0meanPerp = MaterialCLT.g0meanPerp[Grade];
            G0meanInPlane = MaterialCLT.g0meanInPlane[Grade];
            G90mean = MaterialCLT.g90mean[Grade];
            E0_005 = MaterialCLT.e0_005[Grade];
            E90_005 = MaterialCLT.e90_005[Grade];
            G0_005Perp = MaterialCLT.g0_005_Perp[Grade];
            G0_005InPlane = MaterialCLT.g0_005_InPlane[Grade];
            G90_005 = MaterialCLT.g90_005[Grade];
            RhoMean = MaterialCLT.rhoMean[Grade];
            RhoK = MaterialCLT.rhoK[Grade];
            B1_Vertical = MaterialCLT.b1_Vertical[Grade];
            B1_Vertical_Increased = MaterialCLT.b1_Vertical_Increased[Grade];
            B2_Vertical = MaterialCLT.b2_Vertical[Grade];
            B2_Vertical_Increased = MaterialCLT.b2_Vertical_Increased[Grade];
            B1_Horizontal = MaterialCLT.b1_Horizontal[Grade];
            B1_Horizontal_Increased = MaterialCLT.b1_Horizontal_Increased[Grade];
            B2_Horizontal = MaterialCLT.b2_Horizontal[Grade];
            B2_Horizontal_Increased = MaterialCLT.b2_Horizontal_Increased[Grade];

            E = E0mean;
            G = G0meanPerp;
            Density = RhoMean;
        }
        #endregion

        #region Material properties

        public static TimberType type = TimberType.Softwood;

        //Material properties according to manufacters
        public enum ProducerGrade
        {
            KLH, //ETA ETA-06/0138 of 18.01.2021
            StoraEnso,
        };

        public static readonly Dictionary<string, double> fmk = new Dictionary<string, double>

        {
            {"KLH", 24} //ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> ft0k = new Dictionary<string, double>
        {
            {"KLH", 16.5}//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> ft90k = new Dictionary<string, double>
        {
            {"KLH", 0.12}//ETA ETA-06/0138 of 18.01.2021

        };

        public static readonly Dictionary<string, double> fc0k = new Dictionary<string, double>
        {
            {"KLH", 24}//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> fc90k = new Dictionary<string, double>
        {
            {"KLH", 2.2},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> fvk = new Dictionary<string, double>
        {
            {"KLH", 2.7},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> frk = new Dictionary<string, double>
        {
            {"KLH", 1.2},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> e0mean = new Dictionary<string, double>
        {
            {"KLH", 12000},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> e90mean = new Dictionary<string, double>
        {
            {"KLH", 450},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> g0meanPerp = new Dictionary<string, double>
        {
            {"KLH", 690},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> g0meanInPlane = new Dictionary<string, double>
        {
            {"KLH", 500},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> g90mean = new Dictionary<string, double>
        {
            {"KLH", 90},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> e0_005 = new Dictionary<string, double>
        {
            //The ETAs do not provide characteristic values for stiffnesses.
           
            {"KLH", 7400}, //No value provided, Reference C24 used
        };

        public static readonly Dictionary<string, double> e90_005 = new Dictionary<string, double>
        {
            {"KLH", 250}, //No value provided, Reference GL24h used
        };

        public static readonly Dictionary<string, double> g0_005_Perp = new Dictionary<string, double>
        {

            {"KLH", 690 * 2/3}, //According to DIN EN 1995-1-1/NA §NCI Zu 3.2 "Vollholz" (NA.7)  : G05 = 2/3 * Gmean
        };

        public static readonly Dictionary<string, double> g0_005_InPlane = new Dictionary<string, double>
        {
            {"KLH", 500 * 2/3}, //According to DIN EN 1995-1-1/NA §NCI Zu 3.2 "Vollholz" (NA.7)  : G05 = 2/3 * Gmean
        };

        public static readonly Dictionary<string, double> g90_005 = new Dictionary<string, double>
        {
            {"KLH", 90 * 2/3}, //According to DIN EN 1995-1-1/NA §NCI Zu 3.2 "Vollholz" (NA.7)  : G05 = 2/3 * Gmean
        };

        public static readonly Dictionary<string, double> rhoMean = new Dictionary<string, double>
        {
            {"KLH", 420}, //No value provided, Reference C24 used
        };

        public static readonly Dictionary<string, double> rhoK = new Dictionary<string, double>
        {
            {"KLH", 385},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b1_Vertical = new Dictionary<string, double>
        {
            {"KLH", 0.55},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b1_Vertical_Increased = new Dictionary<string, double>
        {
            {"KLH", 0.65},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b2_Vertical = new Dictionary<string, double>
        {
            {"KLH", 0.8},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b2_Vertical_Increased = new Dictionary<string, double>
        {
            {"KLH", 0.9},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b1_Horizontal = new Dictionary<string, double>
        {
            {"KLH", 0.65},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b1_Horizontal_Increased = new Dictionary<string, double>
        {
            {"KLH", 0.75},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b2_Horizontal = new Dictionary<string, double>
        {
            {"KLH", 1.0},//ETA ETA-06/0138 of 18.01.2021
        };

        public static readonly Dictionary<string, double> b2_Horizontal_Increased = new Dictionary<string, double>
        {
            {"KLH", 1.10},//ETA ETA-06/0138 of 18.01.2021
        };













        public static readonly double b0 = 0.65; //according to DIN EN 1995 - 1 - 2 table 3.1
        public static readonly double bn = 0.80; //according to DIN EN 1995 - 1 - 2 table 3.1





        #endregion
    }
}
