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
    [Description("Generic timber material")]
    public class MaterialTimberGeneric : IMaterial, IMaterialTimber
    {

        #region Object properties

        public string Grade { get; set; }

        public double Density { get; set; }

        public double E { get; set; }

        public double G { get; set; }

        public TimberType Type { get; set; }

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

        public MaterialTimberGeneric(string name, TimberType type, double fmyk, double fmzk, double ft0k, double ft90k, double fc0k, double fc90k, double fvk, double frk, double e0mean,
            double e90mean, double g0mean, double e0_005, double g0_005, double rhoMean, double rhoK)
        {
            Grade = name;
            Type = type;
            Fmyk = fmyk;
            Fmzk = fmzk;
            Ft0k = ft0k;
            Ft90k = ft90k;
            Fc0k = fc0k;
            Fc90k = fc90k;
            Fvk = fvk;
            Frk = frk;
            E0mean = e0mean;
            E90mean = e90mean;
            G0mean = g0mean;
            E0_005 = e0_005;
            G0_005 = g0_005;
            RhoMean = rhoMean;
            RhoK = rhoK;

            Density = RhoMean;
            E = E0mean;
            G = G0mean;

        }

        #endregion

    }
}
