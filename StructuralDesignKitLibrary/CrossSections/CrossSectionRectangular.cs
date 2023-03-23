// "This file is part of the 'StructuralDesignKit_Holz (SDK_H)' which is released under the MIT License 
// The project is supported and maintained by https://www.knippershelbig.com/ 
// For more information regarding the porject: https://github.com/knippershelbig/StructuralDesignKit_Holz 
// For the full license details: https://github.com/knippershelbig/StructuralDesignKit_Holz/blob/master/LICENSE"

using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using System;
using System.ComponentModel;

namespace StructuralDesignKitLibrary.CrossSections
{
    public class CrossSectionRectangular : ICrossSection
    {


        #region properties
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public IMaterial Material { get; set; }
        public double Area { get; set; }
        public double MomentOfInertia_Y { get; set; }
        public double MomentOfInertia_Z { get; set; }
        public double TorsionalInertia { get; set; }
        public double SectionModulus_Y { get; set; }
        public double SectionModulus_Z { get; set; }
        public double TorsionalModulus { get; set; }
        public double EIy { get; set; }


        /// <summary>
        /// Width
        /// </summary>
        [Description("Width")]
        private int _B;
        public int B
        {
            get { return _B; }
            set
            {
                if (value <= 0) throw new ArgumentOutOfRangeException("the width of a cross section cannot be inferior or equal to 0");
                else _B = value;
            }
        }

        /// <summary>
        /// Height
        /// </summary>
        [Description("Height")]
        private int _H;
        public int H
        {
            get { return _H; }
            set
            {
                if (value <= 0) throw new ArgumentOutOfRangeException("the height of a cross section cannot be inferior or equal to 0");
                else _H = value;

            }
        }


        #endregion


        #region Constructors 
        /// <summary>
        /// Minimal constructor for rectangular cross section
        /// </summary>
        /// <param name="b">width</param>
        /// <param name="h">height</param>
        /// <param name="material">Material</param>
        public CrossSectionRectangular(int b, int h, IMaterial material)
        {
            B = b;
            H = h;
            if (material is MaterialTimberBaubuche)
            {
                MaterialTimberBaubuche baubuche = (MaterialTimberBaubuche)material;
                baubuche.UpdateBaubucheProperties(b, h);
                Material = baubuche;
            }
            else Material = material;
            this.ComputeCrossSectionProperties();
        }


        /// <summary>
        /// Full definition of the cross section
        /// </summary>
        /// <param name="id">Unique ID</param>
        /// <param name="name">Cross section name</param>
        /// <param name="b">width</param>
        /// <param name="h">height</param>
        /// <param name="material">Material</param>
        public CrossSectionRectangular(int id, string name, int b, int h, IMaterial material) : this(b, h, material)
        {
            ID = id;
            Name = name;
        }
        #endregion

        public void ComputeCrossSectionProperties()
        {
            Area = B * H;
            MomentOfInertia_Y = B * Math.Pow(H, 3) / 12;
            MomentOfInertia_Z = H * Math.Pow(B, 3) / 12;
            SectionModulus_Y = B * Math.Pow(H, 2) / 6;
            SectionModulus_Z = H * Math.Pow(B, 2) / 6;

            //Torsion
            double c1 = (1.0 / 3.0) * (1 - (0.63 / (H / (double)B)) + (0.052 / Math.Pow((H / (double)B), 5)));
            TorsionalInertia = c1 * H * Math.Pow(B, 3);

            double c2 = 1 - (0.65 / (1 + Math.Pow((H / (double)B), 3)));
            TorsionalModulus = (c1 / c2) * H * Math.Pow(B, 2);
            EIy = Material.E * MomentOfInertia_Y;

        }

        #region Compute Stresses
        public double ComputeStressBendingY(double BendingMomentY)
        {
            return Math.Abs(BendingMomentY) * 1e6 / SectionModulus_Y;
        }

        public double ComputeStressBendingZ(double BendingMomentZ)
        {
            return Math.Abs(BendingMomentZ) * 1e6 / SectionModulus_Z;
        }

        public double ComputeNormalStress(double NormalForce)
        {
            return Math.Abs(NormalForce) * 1e3 / Area;
        }

        public double ComputeShearY(double ShearForceY)
        {
            return Math.Abs(ShearForceY) * 3 / 2 * 1e3 / Area;
        }

        public double ComputeShearZ(double ShearForceZ)
        {
            return Math.Abs(ShearForceZ) * 3 / 2 * 1e3 / Area;
        }

        public double ComputeTorsion(double TorsionMoment)
        {
            return Math.Abs(TorsionMoment) * 1e6 / TorsionalModulus;
        }

        #endregion


        //Fire design
        public CrossSectionRectangular ComputeReducedCrossSection(int fireDuration, bool top, bool bottom, bool left, bool right)
        {

            double d_ef = EC5.EC5_Utilities.ComputeCharringDepthUnprotected(fireDuration, (IMaterialTimber)this.Material);
            double b = this.B;
            double h = this.H;
            if (left) b -= d_ef;
            if (right) b -= d_ef;
            if (top) h -= d_ef;
            if (bottom) h -= d_ef;

            return new CrossSectionRectangular((Int32)Math.Floor(b), (Int32)Math.Floor(h), this.Material);

        }
    }
}
