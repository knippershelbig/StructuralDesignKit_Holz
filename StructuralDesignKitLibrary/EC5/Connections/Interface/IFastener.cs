using StructuralDesignKitLibrary.EC5.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.Interface
{
    public interface IFastener
    {
        /// <summary>
        /// Type of fasterner (Bolt, dowels, screw...)
        /// </summary>
        [Description("Type of fasterner (Bolt, dowels, screw...)")]
        EC5.EC5_Utilities.FastenerType Type { get; }

        /// <summary>
        /// Diameter of the fastener in mm
        /// </summary>
        [Description("Diameter of the fastener")]
        double Diameter { get; set; }

        /// <summary>
        /// Tensile strength of the fasterner in N/mm²
        /// </summary>
        [Description("Tensile strength of the fasterner in N/mm²")]
        double Fuk { get; set; }

        /// <summary>
        /// Characteristic yield moment of the fastener in N.mm
        /// </summary>
        [Description("Characteristic yield moment of the fastener in N.mm")]
        double MyRk { get; set; }


        /// <summary>
        /// Embedment Strength of the fastener in N/mm²
        /// </summary>
        [Description("Embedment Strength of the fastener in N/mm²")]
        double Fhk { get; set; }

        /// <summary>
        /// Withdrawal strength of the fastener in N
        /// </summary>
        [Description("Withdrawal strength of the fastener in N")]
        double WithdrawalStrength { get; set; }

        /// <summary>
        /// Limitation of the Johansen part for the rope effect based on EN 1995-1-1 §8.2.2 (2)
        /// </summary>
        [Description("Limitation of the Johansen part for the rope effect based on EN 1995-1-1 §8.2.2 (2)")]
        double MaxJohansenPart { get; set; }

        /// <summary>
        /// Minimum spacing parallel to grain in mm
        /// </summary>
        [Description("Minimum spacing parallel to grain in mm")]
        double a1min { get; set; }

        /// <summary>
        /// Minimum spacing perpendicular to grain in mm
        /// </summary>
        [Description("Minimum spacing perpendicular to grain in mm")]
        double a2min { get; set; }

        /// <summary>
        /// Minimum spacing to loaded end in mm
        /// </summary>
        [Description("Minimum spacing to loaded end in mm")]
        double a3tmin { get; set; }

        /// <summary>
        /// Minimum spacing to unloaded end in mm
        /// </summary>
        [Description("Minimum spacing to unloaded end in mm")]
        double a3cmin { get; set; }

        /// <summary>
        /// Minimum spacing to loaded edge in mm
        /// </summary>
        [Description("Minimum spacing to loaded edge in mm")]
        double a4tmin { get; set; }

        /// <summary> 
        /// Minimum spacing to unloaded edge in mm
        /// </summary>
        [Description("Minimum spacing to unloaded edge in mm")]
        double a4cmin { get; set; }

        /// <summary>
        /// Computes the effective number of fasteners
        /// The spacing parallel to the grain a1 influences the effective number of fasteners and hence the overall load‐bearing capacity of the joint, since an insufficient
        /// spacing a1 may trigger premature splitting of the joint along the set of fasteners in the grain direction
        /// </summary>
        /// <param name="n">number of fasteners aligned parallel to the grain</param>
        /// <param name="a1">spacing parallel to the grain</param>
        /// <param name="angle">Load angle toward the timber grain in degree</param>
        /// <returns></returns>
        [Description("Computes the effective number of fasteners")]
        double ComputeEffectiveNumberOfFastener(int n, double a1, double angle);

        /// <summary>
        /// Computes the embedment strength of the fastener
        /// </summary>
        /// <param name="timber">Timber connected with the fastener</param>
        /// <param name="angle">Load angle toward the timber grain in degree</param>
        /// <returns></returns>
        void ComputeEmbedmentStrength(IMaterialTimber timber, double angle);

        void ComputeWithdrawalStrength(IShearCapacity ConnectionType);

        void ComputeSpacings(double angle);



    }
}
