using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.CrossSections.Interfaces
{
    public interface iCrossSection

    {
        /// <summary>
        /// ID - Unique identifier
        /// </summary>
        [Description("ID")]
        int ID { get; set; }


        /// <summary>
        /// Name
        /// </summary>
        [Description("Name")]
        string Name { get; set; }

        /// <summary>
        /// Additional description
        /// </summary>
        [Description("Additional description")]
        string Description { get; set; }


        /// <summary>
        /// Material Object
        /// </summary>
        [Description("Material Object")]
        IMaterial Material { get; set; }

        /// <summary>
        /// Cross section's area in mm²
        /// </summary>
        [Description("Cross section's area in mm²")]
        double Area { get; set; }


        /// <summary>
        /// Cross section's Moment of Inertia along Y [mm4]
        /// </summary>
        [Description("Cross section's Moment of Inertia along Y [mm4]")]
        double MomentOfInertia_Y { get; set; }

        /// <summary>
        /// Cross section's Moment of Inertia along Z [mm4]
        /// </summary>
        [Description("Cross section's Moment of Inertia along Z [mm4]")]
        double MomentOfInertia_Z { get; set; }

        /// <summary>
        /// Section Modulus Y [mm3]
        /// </summary>
        [Description("Section Modulus Y [mm3]")]
        double SectionModulus_Y { get; set; }

        /// <summary>
        /// Section Modulus Z [mm3]
        /// </summary>
        [Description("Section Modulus Z [mm3]")]
        double SectionModulus_Z { get; set; }

        /// <summary>
        /// Computes the cross section properties given the basic defining parameters
        /// </summary>
        [Description("Computes the cross section properties given the basic defining parameters")]
        void ComputeCrossSectionProperties();
    }
}
