using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace StructuralDesignKitLibrary.Connections.Interface
{
    public interface ISteelTimberShear
    {
        /// <summary>
        /// Fastener used in the shear plane considered
        /// </summary>
        IFastener Fastener { get; set; }

        /// <summary>
        /// Thickness of the steel plate
        /// </summary>
        double SteelPlateThickness { get; set; }


        /// <summary>
        /// Boolean value which defines if the steel plate is considered as thick according to EN 1995-1-1 §8.2.3 (1)
        /// </summary>
        bool isTkickPlate { get; set; }

        /// <summary>
        /// Boolean value which defines if the steel plate is considered as thin according to EN 1995-1-1 §8.2.3 (1)
        /// </summary>
        bool isThinPlate { get; set; }


        /// <summary>
        /// Load angle toward the timber grain in degree
        /// </summary>
        double Angle { get; set; }


        /// <summary>
        /// Timber material used in the shear plane
        /// </summary>
        IMaterialTimber Timber { get; set; }


        /// <summary>
        /// Timber thickness considered in the shear plane (t1 or t2 according to 8.2.3 (3))
        /// </summary>
        double TimberThickness { get; set; }


        /// <summary>
        /// List of the failure mechanisms 
        /// </summary>
        List<string> FailureModes { get; set; }


        /// <summary>
        /// List of the failure mechanism capacities
        /// </summary>
        List<double> Capacities { get; set; }

        /// <summary>
        /// Driving failure mode
        /// </summary>
        string FailureMode { get; set; }

        /// <summary>
        /// Capacity of the fastener. This value is the minimum of the failure modes multiplied by the number of shear plane considered for the failure mode
        /// </summary>
        double Capacity { get; set; }

        /// <summary>
        /// Computes all the relevant failure modes
        /// </summary>
        void ComputeFailingModes();


        /// <summary>
        /// Boolean value which defines if the rope effect should be considered
        /// </summary>
        bool RopeEffect { get; set; }




        /*
         SteelSingleOuterPlate
         SteelDoubleOuterPlates
         SteelSingleInnerplate
         SteelMultipleInnerPlates
        */
    }
}
