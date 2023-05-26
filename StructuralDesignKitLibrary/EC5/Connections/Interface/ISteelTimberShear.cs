using StructuralDesignKitLibrary.EC5.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StructuralDesignKitLibrary.Connections.Interface
{
    public interface ISteelTimberShear: IShearCapacity
    {

        /// <summary>
        /// Thickness of the steel plate
        /// </summary>
        [Description("Thickness of the steel plate")]
        double SteelPlateThickness { get; set; }


        /// <summary>
        /// Boolean value which defines if the steel plate is considered as thick according to EN 1995-1-1 §8.2.3 (1)
        /// </summary>
        [Description("Boolean value which defines if the steel plate is considered as thick according to EN 1995-1-1 §8.2.3 (1)")]
        bool isTkickPlate { get; set; }

        /// <summary>
        /// Boolean value which defines if the steel plate is considered as thin according to EN 1995-1-1 §8.2.3 (1)
        /// </summary>
        [Description("Boolean value which defines if the steel plate is considered as thin according to EN 1995-1-1 §8.2.3 (1)")]
        bool isThinPlate { get; set; }


        /// <summary>
        /// Load angle toward the timber grain in degree
        /// </summary>
        [Description("Load angle toward the timber grain in degree")]
        double Angle { get; set; }


        /// <summary>
        /// Timber material used in the shear plane
        /// </summary>
        [Description("Timber material used in the shear plane")]
        IMaterialTimber Timber { get; set; }


        /// <summary>
        /// Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))
        /// </summary>
        [Description("Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))")]
        double TimberThickness { get; set; }


        /*
        Remaining to do
         SteelMultipleInnerPlates
        */
    }
}
