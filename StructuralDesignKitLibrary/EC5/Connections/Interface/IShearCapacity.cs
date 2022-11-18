using StructuralDesignKitLibrary.Connections.Interface;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.EC5.Connections.Interface
{
    public interface IShearCapacity
    {
        /// <summary>
        /// Fastener used in the shear plane considered
        /// </summary>
        [Description("Fastener used in the shear plane considered")]
        IFastener Fastener { get; set; }

        /// <summary>
        /// List of the failure mechanisms 
        /// </summary>
        [Description("List of the failure mechanisms")]
        List<string> FailureModes { get; set; }


        /// <summary>
        /// List of the failure mechanism capacities
        /// </summary>
        [Description("List of the failure mechanism capacities")]
        List<double> Capacities { get; set; }

        /// <summary>
        /// Driving failure mode
        /// </summary>
        [Description("Driving failure mode")]
        string FailureMode { get; set; }

        /// <summary>
        /// Capacity of the fastener. This value is the minimum of the failure modes multiplied by the number of shear plane considered for the failure mode
        /// </summary>
        [Description("apacity of the fastener. This value is the minimum of the failure modes multiplied by the number of shear plane considered for the failure mode")]
        double Capacity { get; set; }

        /// <summary>
        /// Computes all the relevant failure modes
        /// </summary>
        [Description("Computes all the relevant failure modes")]
        void ComputeFailingModes();


        /// <summary>
        /// Boolean value which defines if the rope effect should be considered
        /// </summary>
        [Description("Boolean value which defines if the rope effect should be considered")]
        bool RopeEffect { get; set; }
    }
}
