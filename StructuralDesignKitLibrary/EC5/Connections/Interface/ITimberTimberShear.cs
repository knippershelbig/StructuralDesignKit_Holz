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
    public interface ITimberTimberShear:IShearCapacity
    {

        /// <summary>
        /// Load angle toward the timber 1 grain in degree
        /// </summary>
        [Description("Load angle toward the timber 1 grain in degree")]
        double Angle1 { get; set; }

        /// <summary>
        /// Load angle toward the timber 2 grain in degree
        /// </summary>
        [Description("Load angle toward the timber 2 grain in degree")]
        double Angle2 { get; set; }


        /// <summary>
        /// First timber material used in the shear plane
        /// </summary>
        [Description("First timber material used in the shear plane")]
        IMaterialTimber Timber1 { get; set; }

        /// <summary>
        /// Second timber material used in the shear plane
        /// </summary>
        [Description("Second timber material used in the shear plane")]
        IMaterialTimber Timber2 { get; set; }


        /// <summary>
        /// Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))
        /// </summary>
        [Description("Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))")]
        double T1 { get; set; }

        /// <summary>
        /// Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))
        /// </summary>
        [Description("Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))")]

        double T2 { get; set; }

        /// <summary>
        /// Embedment strength of the fastener on timber 1
        /// </summary>
        [Description("Embedment strength of the fastener on timber 1")]
        double Fhk1 { get; set; }

        /// <summary>
        /// Embedment strength of the fastener on timber 1
        /// </summary>
        [Description("Embedment strength of the fastener on timber 1")]
        double Fhk2 { get; set; }

       




        /*
TimberTimberSingleShear
TimberTimberDoubleShear
        */
    }
}
