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

			//List<double> thicknesses = new List<double>() { 40, 20, 20, 20, 40 };
			//List<int> orientations = new List<int>() { 0, 90, 0, 90, 0 };
			List<double> thicknesses = new List<double>() { 40,40,40,40,40 };
			List<int> orientations = new List<int>() { 0, 90, 0, 90, 0 };
			MaterialTimberSoftwood c24 = new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24);
			MaterialTimberGlulam Glulam = new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h);
			MaterialTimberGeneric t1 = new MaterialTimberGeneric(new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24), new List<string>() { "E0mean" }, new List<object>() { 11000 }, "toto");
			MaterialTimberGeneric t2 = new MaterialTimberGeneric(new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C16), new List<string>() { "E0mean" }, new List<object>() { 8000 }, "tata");
			MaterialTimberGeneric t3 = new MaterialTimberGeneric(new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C30), new List<string>() { "E0mean" }, new List<object>() { 12000 }, "tata");
			MaterialTimberGeneric LVL = new MaterialTimberGeneric(new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C30), new List<string>() { "E0mean" }, new List<object>() { 10500 }, "tata");
			List<IMaterialTimber> materialTimbers = new List<IMaterialTimber>();

			for (int i = 0; i < thicknesses.Count; i++)
			{
				materialTimbers.Add(c24);
			}



			///CLT DEV
			CLT_Layup CLTLayup = new CLT_Layup(thicknesses, orientations, materialTimbers);

			Console.WriteLine("AreaX = {0:0}\nAreaY = {1:0}\nCoGX = {2:0}\nCOGY={3:0}\nZoXX={4:0}\nZuXX={5:0}\nZoYY={6:0}\nZuYY={7:0}\n" +
				"InertiaX={8:0}\nInertiaY={9:0}\nSectionMod_X={10:0}\nSectionMod_Y={11:0}" +
				"\nStaticMoment_R_X ={12:0}\nStaticMoment_R_Y ={13:0}\nStaticMoment_X ={14:0}\nStaticMoment_X ={15:0}",
				CLTLayup.CS_X.Area,
				CLTLayup.CS_Y.Area,
				CLTLayup.CS_X.CenterOfGravity,
				CLTLayup.CS_Y.CenterOfGravity,
				CLTLayup.CS_X.Zo,
				CLTLayup.CS_X.Zu,
				CLTLayup.CS_Y.Zo,
				CLTLayup.CS_Y.Zu,
				CLTLayup.CS_X.MomentOfInertia,
				CLTLayup.CS_Y.MomentOfInertia,
				CLTLayup.CS_X.W0_Net,
				CLTLayup.CS_Y.W0_Net,
				CLTLayup.CS_X.Sr0Net,
				CLTLayup.CS_Y.Sr0Net,
				CLTLayup.CS_X.S0Net,
				CLTLayup.CS_Y.S0Net);


			Console.WriteLine("Capacities\n" +
				"Mx={0:0.0}\nMy={1:0.0}\nN_compX={2:0.0}\nN_compY={3:0.0}\n\nN_TensX={4:0.0}\nN_TensY={5:0.0}\nVx={6:0.0}\nVy={7:0.0}",
				CLTLayup.CS_X.MChar,
				CLTLayup.CS_Y.MChar,
				CLTLayup.CS_X.NCompressionChar,
				CLTLayup.CS_Y.NCompressionChar,
				CLTLayup.CS_X.NTensionChar * 0.8 / (1.3 * 1.5),
				CLTLayup.CS_Y.NTensionChar,
				CLTLayup.CS_X.VChar,
				CLTLayup.CS_Y.VChar);
			///CLT DEV

			double charLoad = CLTLayup.CS_X.NTensionChar * 0.8 / (1.3 * 1.5);
			CLTLayup.CS_X.ComputeNormalStress(charLoad * 1.5);

			Console.WriteLine("Normal stresses in layers");
			Console.WriteLine(" Char Load considered {0:0}kN", charLoad);
			foreach (double stress in CLTLayup.CS_X.SigmaNTens)
			{
				Console.WriteLine(stress.ToString());
			}




			Console.WriteLine("\n\nBENDING STRESSES\n\n");

			CLTLayup.CS_X.ComputeBendingStress(CLTLayup.CS_X.MChar*0.8/1.3);

			for (int i = 0; i < CLTLayup.CS_X.LamellaOrientations.Count; i++)
			{

				Console.WriteLine("Layer {0}", i);
				Console.WriteLine("Top: {0:0.0}N/mm²\nMiddle: {1:0.0}N/mm²\nBottom: {2:0.0}N/mm²\n",
					CLTLayup.CS_X.SigmaM[i][0],
					CLTLayup.CS_X.SigmaM[i][1],
					CLTLayup.CS_X.SigmaM[i][2]);


			}

			Console.WriteLine("\nMax Bending : {0:0.0}", CLTLayup.CS_X.MChar);

			Console.ReadLine();
		}
	}

}
