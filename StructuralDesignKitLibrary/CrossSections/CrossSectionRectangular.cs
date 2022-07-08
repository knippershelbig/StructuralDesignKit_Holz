using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.CrossSections
{
    public class CrossSectionRectangular : iCrossSection
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public IMaterial Material { get; set; }
        public double Area { get; set; }
        public double MomentOfInertia_Y { get; set; }
        public double MomentOfInertia_Z { get; set; }
        public double SectionModulus_Y { get; set; }
        public double SectionModulus_Z { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        [Description("Width")]
        public int B { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        [Description("Height")]
        public int H { get; set; }


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
            Material = material;
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


        public void ComputeCrossSectionProperties()
        {
            Area = B * H;
            MomentOfInertia_Y = B * Math.Pow(H, 3) / 12;
            MomentOfInertia_Z = H * Math.Pow(B, 3) / 12;
            SectionModulus_Y = B * Math.Pow(H, 2) / 6;
            SectionModulus_Z = H * Math.Pow(B, 2) / 6;
        }

    }
}
