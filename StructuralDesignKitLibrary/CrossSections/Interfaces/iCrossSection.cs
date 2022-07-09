using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.CrossSections.Interfaces
{
    public interface ICrossSection

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
        /// Cross section's Moment of Inertia in torsion [mm4]
        /// </summary>
        [Description("Cross section's Moment of Inertia in torsion [mm4]")]
        double TorsionalInertia { get; set; }


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
        /// Torsional Section Modulus Wt [mm3]
        /// </summary>
        [Description("Torsional Section Modulus Wt [mm3]")]
        double TorsionalModulus { get; set; }





        /// <summary>
        /// Computes the cross section properties given the basic defining parameters
        /// </summary>
        [Description("Computes the cross section properties given the basic defining parameters")]
        void ComputeCrossSectionProperties();


        /// <summary>
        /// Computes the bending stress on Y
        /// </summary>
        /// <param name="BendingMomentY">Bending moment on Y axis in KN.m</param>
        /// <returns>Noraml stress in N/mm²</returns>
        [Description("Computes the bending stress on Y")]
        double ComputeStressBendingY(double BendingMomentY);

        /// <summary>
        /// Computes the bending stress on Z
        /// </summary>
        /// <param name="BendingMomentZ">Bending moment on Z axis in KN.m</param>
        /// <returns>Noraml stress in N/mm²</returns>
        [Description("Computes the bending stress on Z")]
        double ComputeStressBendingZ(double BendingMomentZ);

        /// <summary>
        /// Computes the normal stress
        /// </summary>
        /// <param name="NormalForce">Normal force in KN</param>
        /// <returns>Noraml stress in N/mm²</returns>
        [Description("Computes the normal stress")]
        double ComputeNormalStress(double NormalForce);

        /// <summary>
        /// Computes the shear stress on Y
        /// </summary>
        /// <param name="ShearForceY">Shear force along Y in KN</param>
        /// <returns>shear stress in N/mm²</returns>
        [Description("Computes the shear stress on Y")]
        double ComputeShearY(double ShearForceY);

        /// <summary>
        /// Computes the shear stress on Z
        /// </summary>
        /// <param name="ShearForceZ">Shear force along Z in KN</param>
        /// <returns>shear stress in N/mm²</returns>
        [Description("Computes the shear stress on Z")]
        double ComputeShearZ(double ShearForceZ);

        /// <summary>
        /// Computes the torsion stress
        /// </summary>
        /// <param name="TorsionMoment">Torsion Moment in KN.m</param>
        /// <returns>Torsion Shear stress in N/mm²</returns>
        /// 
        [Description("Computes the torsion stress")]
        double ComputeTorsion(double TorsionMoment);





    }
}
