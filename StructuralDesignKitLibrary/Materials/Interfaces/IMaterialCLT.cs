using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Materials
{
    public interface IMaterialCLT : IMaterial
    {

        /// <summary>
        /// Timber type according to EN 1995
        /// </summary>
        [Description("Timber type according to EN 1995")]
        EC5.EC5_Utilities.TimberType Type { get; }

        /// <summary>
        /// Characteristic bending strength - Y axis
        /// </summary>
        [Description("Characteristic bending strength")]
        double Fmk { get; set; }

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
        /// Mean shear Modulus parallel to the grain
        /// </summary>
        [Description("Mean shear Modulus parallel to the grain | Perpendicular to the solid wood slab ")]
        double G0meanPerp { get; set; }

        /// <summary>
        /// Mean shear Modulus parallel to the grain
        /// </summary>
        [Description("Mean shear Modulus parallel to the grain | In plane of the solid wood slab ")]
        double G0meanInPlane { get; set; }

        /// <summary>
        /// Mean shear Modulus parallel to the grain
        /// </summary>
        [Description("Mean rolling shear Modulus parallel to the grain")]
        double G90mean { get; set; }

        /// <summary>
        /// Characteristic Modulus of Elasticity parallel to grain
        /// </summary>
        [Description("Characteristic Modulus of Elasticity parallel to grain")]
        double E0_005 { get; set; }

        /// <summary>
        /// Characteristic Modulus of Elasticity perpendicular to grain
        /// </summary>
        [Description("Characteristic Modulus of Elasticity perpendicular to grain")]
        double E90_005 { get; set; }

        /// <summary>
        /// Characteristic shear Modulus parallel to the grain
        /// </summary>
        [Description("Characteristic shear Modulus parallel to the grain | Perpendicular to the solid wood slab ")]
        double G0_005Perp { get; set; }

        /// <summary>
        /// Characteristic shear Modulus parallel to the grain
        /// </summary>
        [Description("Characteristic shear Modulus parallel to the grain | In plane of the solid wood slab ")]
        double G0_005InPlane { get; set; }

        /// <summary>
        /// Characteristic shear Modulus parallel to the grain
        /// </summary>
        [Description("Characteristic rolling shear Modulus parallel to the grain")]
        double G90_005 { get; set; }

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

        /// <summary>
        /// Regular charring rate within one single layer - Vertical elements
        /// </summary>
        [Description("Regular charring rate within one single layer - Vertical elements")]
        double B1_Vertical { get; set; }

        /// <summary>
        /// Charring rate within one single layer for local effects - Increased value for a solid wood slabs with width b < 300 mm  - Vertical elements
        /// </summary>
        [Description("Charring rate within one single layer for local effects - Increased value for a solid wood slabs with width b < 300 mm - Vertical elements")]
        double B1_Vertical_Increased { get; set; }

        /// <summary>
        /// Increased charring rate after the failure / drop off of one layer  - Vertical elements
        /// </summary>
        [Description("Increased charring rate after the failure / drop off of one layer  - Vertical elements")]
        double B2_Vertical { get; set; }

        /// <summary>
        /// Increased charring rate after the failure / drop off of one layer - Increased value for a solid wood slabs with width b < 300 mm  - Vertical elements
        /// </summary>
        [Description("Increased charring rate after the failure / drop off of one layer - Increased value for a solid wood slabs with width b < 300 mm - Vertical elements")]
        double B2_Vertical_Increased { get; set; }

        /// <summary>
        /// Regular charring rate within one single layer - Horizontal elements
        /// </summary>
        [Description("Regular charring rate within one single layer - Horizontal elements")]
        double B1_Horizontal { get; set; }

        /// <summary>
        /// Charring rate within one single layer for local effects - Increased value for a solid wood slabs with width b < 300 mm  - Horizontal elements
        /// </summary>
        [Description("Charring rate within one single layer for local effects - Increased value for a solid wood slabs with width b < 300 mm - Horizontal elements")]
        double B1_Horizontal_Increased { get; set; }

        /// <summary>
        /// Increased charring rate after the failure / drop off of one layer  - Horizontal elements
        /// </summary>
        [Description("Increased charring rate after the failure / drop off of one layer  - Horizontal elements")]
        double B2_Horizontal { get; set; }

        /// <summary>
        /// Increased charring rate after the failure / drop off of one layer - Increased value for a solid wood slabs with width b < 300 mm  - Horizontal elements
        /// </summary>
        [Description("Increased charring rate after the failure / drop off of one layer - Increased value for a solid wood slabs with width b < 300 mm - Horizontal elements")]
        double B2_Horizontal_Increased { get; set; }
    }
}
