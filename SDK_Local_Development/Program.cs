using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StructuralDesignKitLibrary;
using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.Materials;

namespace SDK_Local_Development
{
	internal class Program
	{
		static void Main(string[] args)
		{
			List<int> thicknesses = new List<int>() { 40,20,40};

			IMaterialTimber c24 = new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24);
			List<IMaterialTimber> materials = new List<IMaterialTimber>() { c24, c24, c24 };
			List<int> orientations = new List<int>() { 0, 90, 0 };

			CrossSectionCLT CLT_CS = new CrossSectionCLT(thicknesses,orientations,materials);

			for (int i = 0; i < CLT_CS.NbOfLayers; i++)
			{
                Console.WriteLine("Layer {0}: {1}mm - {2}° - {3}", i,
					CLT_CS.LamellaThicknesses[i], CLT_CS.LamellaOrientations[i], CLT_CS.LamellaMaterials[i].Grade.ToString());
            }

            Console.WriteLine("CLT thickness = {0:0}mm", CLT_CS.Thickness);
            Console.WriteLine("COG_X = {0:0}mm\nCOG_Y={1:0}mm",CLT_CS.CenterOfGravityXX, CLT_CS.CenterOfGravityYY);
            Console.WriteLine("AreaX = {0:0}mm² -- AreaY={1:0}mm",CLT_CS.AreaX, CLT_CS.AreaY);
            Console.WriteLine("Ixx = {0:0}mm4 -- Iyy={1:0}mm",CLT_CS.MomentOfInertiaXX, CLT_CS.MomentOfInertiaYY);
			Console.ReadLine();
        }
	}
}
