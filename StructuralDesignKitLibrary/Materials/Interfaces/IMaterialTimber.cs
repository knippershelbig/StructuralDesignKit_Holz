using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Materials
{
    internal interface IMaterialTimber:IMaterial    
    {

        /// <summary>
        /// Timber type according to EN 1995
        /// </summary>
        [Description("Timber type according to EN 1995")]
        EC5.EC5_Utilities.TimberType Type { get; }

        /// <summary>
        /// Characteristic bending strength - Y axis
        /// </summary>
        [Description("Characteristic bending strength - Y axis")]
        double Fmyk { get; set; }

        /// <summary>
        /// Characteristic bending strength - Z axis
        /// </summary>
        [Description("Characteristic bending strength - Z axis")]
        double Fmzk { get; set; }

        /// <summary>
        /// Characteristic tension strength parallel to grain
        /// </summary>
        [Description("Characteristic tension strength parallel to grain")]
        double Ft0k { get; set; }

        /// <summary>
        /// Characteristic tension strength perpendicular to grain
        /// </summary>
        [Description("Characteristic tension strength perpendicular to grain")]
        double Ft90k { get; set; }

        /// <summary>
        /// Characteristic compression strength parallel to grain
        /// </summary>
        [Description("Characteristic compression strength parallel to grain")]
        double Fc0k { get; set; }

        /// <summary>
        /// Characteristic compression strength perpendicular to grain
        /// </summary>
        [Description("Characteristic compression strength perpendicular to grain")]
        double Fc90k { get; set; }


        /// <summary>
        /// Characteristic shear strength parallel to grain
        /// </summary>
        [Description("Characteristic shear strength parallel to grain")]
        double Fvk { get; set; }


        /// <summary>
        /// Characteristic rolling shear strength
        /// </summary>
        [Description("Characteristic rolling shear strength")]
        double Frk { get; set; }

        /// <summary>
        /// Mean Modulus of Elasticity parallel to grain
        /// </summary>
        [Description("Mean Modulus of Elasticity parallel to grain")]
        double E0mean { get; set; }



        /// <summary>
        /// Mean Modulus of Elasticity perpendicular to grain
        /// </summary>
        [Description("Mean Modulus of Elasticity perpendicular to grain")]
        double E90mean { get; set; }


        /// <summary>
        /// Mean shear Modulus
        /// </summary>
        [Description("Mean shear Modulus")]
        double G0mean { get; set; }

        /// <summary>
        /// Characteristic Modulus of Elasticity parallel to grain
        /// </summary>
        [Description("Characteristic Modulus of Elasticity parallel to grain")]
        double E0_005 { get; set; }


        /// <summary>
        /// Characteristic shear Modulus
        /// </summary>
        [Description("Characteristic shear Modulus")]
        double G0_005 { get; set; }

        /// <summary>
        /// Mean Density
        /// </summary>
        [Description("Mean Density")]
        double RhoMean { get; set; }


        /// <summary>
        /// Characteristic Density
        /// </summary>
        [Description("Characteristic Density")]
        double RhoK { get; set; }


    }
}
