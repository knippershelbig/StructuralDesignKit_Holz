using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using StructuralDesignKitLibrary.EC5;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;
using System.Reflection;

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

        public double B0 { get; set; }

        public double Bn { get; set; }

        #endregion

        #region constructor

        public MaterialTimberGeneric(string name, TimberType type, double fmyk, double fmzk, double ft0k, double ft90k, double fc0k, double fc90k, double fvk, double frk, double e0mean,
            double e90mean, double g0mean, double e0_005, double g0_005, double rhoMean, double rhoK, double b0, double bn)
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
            B0=b0;  
            Bn= bn; 

            Density = RhoMean;
            E = E0mean;
            G = G0mean;
        }

        public MaterialTimberGeneric(IMaterialTimber baseMaterial, List<string> propertiesToModify, List<object> values, string tag)
        {

            //Make sure both list have the same length and their content is not null
            if (propertiesToModify == null) throw new Exception("The propertiesToModify list is null");
            if (values == null) throw new Exception("The propertiesToModify list is null");

            if (propertiesToModify.Count != values.Count) throw new Exception("The propertiesToModify list and the values list do not have the same length");

            //To refactor
            this.Grade = tag;
            //

            this.Type = baseMaterial.Type;
            this.Fmyk = baseMaterial.Fmyk;
            this.Fmzk = baseMaterial.Fmzk;
            this.Ft0k = baseMaterial.Ft0k;
            this.Ft90k = baseMaterial.Ft90k;
            this.Fc0k = baseMaterial.Fc0k;
            this.Fc90k = baseMaterial.Fc90k;
            this.Fvk = baseMaterial.Fvk;
            this.Frk = baseMaterial.Frk;
            this.E0mean = baseMaterial.E0mean;
            this.E90mean = baseMaterial.E90mean;
            this.G0mean = baseMaterial.G0mean;
            this.E0_005 = baseMaterial.E0_005;
            this.G0_005 = baseMaterial.G0_005;
            this.RhoMean = baseMaterial.RhoMean;
            this.RhoK = baseMaterial.RhoK;
            this.B0 = baseMaterial.B0;
            this.Bn = baseMaterial.Bn;

            this.Density = baseMaterial.Density;
            this.E = baseMaterial.E;
            this.G = baseMaterial.G;

            var typ = typeof(IMaterialTimber);
            List<PropertyInfo> properties = typeof(IMaterialTimber).GetProperties().ToList();
            properties.AddRange(typeof(IMaterial).GetProperties().ToList());


            int count = 0;
            foreach (string property in propertiesToModify)

            {
                if (properties.Any(p => p.Name == property))
                {
                    var prop = this.GetType().GetProperty(property);
                    //Get property type:
                    var propType = prop.PropertyType;
                    prop.SetValue(this, Convert.ChangeType(values[count], propType, null));
                }
                else throw new Exception(String.Format("The property \"{0}\" does not exist", property));
                count += 1;
            }
        }

        #endregion

    }
}
