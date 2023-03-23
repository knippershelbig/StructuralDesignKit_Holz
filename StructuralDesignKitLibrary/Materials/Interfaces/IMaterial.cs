using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Materials
{
    [Description("Interface for materials - regroups the basic properties needed to define a material")]
    public interface IMaterial
    {
        /// <summary>
        /// Material grade according to the relevant standard
        /// </summary>
        [Description("Material grade according to the relevant standard")]
        string Grade { get; set; }

        /// <summary>
        /// Density used for mass calculation in kg/m³
        /// </summary>
        [Description("Density used for mass calculation")]
        double Density { get; set; }

        /// <summary>
        /// Young Modulus
        /// </summary>
        [Description("Young Modulus")]
        double E { get; set; }

        /// <summary>
        /// Shear Modulus
        /// </summary>
        [Description("Shear Modulus")]
        double G { get; set; }
    }
}

