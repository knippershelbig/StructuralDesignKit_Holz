using Dlubal.RFEM5;
using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.CrossSections
{
	public class CrossSectionCLT
	{
		public int ID { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public int Thickness { get; set; }
		public int NbOfLayers { get; set; }

		/// <summary>
		/// List of lamella thicknesses
		/// </summary>
		public List<int> LamellaThicknesses { get; set; }

		/// <summary>
		/// List of lamella orientation (0° or 90°)
		/// </summary>
		private List<int> _orientations;

		public List<int> LamellaOrientations
		{
			get { return _orientations; }
			set
			{
				//Verify if the lamella orientation is properly defined
				foreach (int angle in value)
				{
					if (angle != 0 && angle != 90) throw new Exception("The CLT lamella angle can only be 0° or 90°");
				}
				_orientations = value;
			}
		}

		/// <summary>
		/// List of the material constituing each lamella
		/// </summary>
		public List<IMaterialTimber> LamellaMaterials { get; set; }

		/// <summary>
		/// Position of the lamella center of gravity from the CS upper edge
		/// </summary>
		public List<double> LamellaDistanceFromTopFibre { get; set; }

		/// <summary>
		/// Distance from the center of gravity of a layer to overall center of gravity
		/// </summary>
		public List<double> LamellaDistanceToCDG { get; set; }

		/// <summary>
		/// define the with of a single lamella
		/// </summary>
		public int LamellaWidth { get; set; }

		/// <summary>
		/// define if the narrow side of the lamella is glued
		/// </summary>
		public bool NarrowSideGlued { get; set; }

		//Reference E modulus in the X axis to determine the geometrical properties
		private double EcX { get; set; }
		private double EcY { get; set; }

        private int FirstLayer0{ get; set; }
        private int FirstLayer90{ get; set; }
        private int LastLayer0{ get; set; }
        private int LastLayer90{ get; set; }


		public double CenterOfGravityXX { get; set; }
		public double CenterOfGravityYY { get; set; }

        public double ZuXX{ get; set; }
        public double ZoXX{ get; set; }
        public double ZuYY{ get; set; }
        public double ZoYY{ get; set; }
		
		public double AreaX { get; set; }
		public double AreaY { get; set; }
		public double MomentOfInertiaXX { get; set; }
		public double MomentOfInertiaYY { get; set; }
		/// <summary>
		/// Section Module 0°
		/// </summary>
		public double WXX { get; set; }

		/// <summary>
		/// Section Module 90°
		/// </summary>
		public double WYY { get; set; }


		public double TorsionalInertia { get; set; }
		public double SectionModulus_XX { get; set; }
		public double SectionModulus_YY { get; set; }
		public double TorsionalModulus { get; set; }
		public double NxCompressionChar { get; set; }
		public double NxTensionChar { get; set; }
		public double NyCompressionChar { get; set; }
		public double NyTensionChar { get; set; }
		public double NxyChar { get; set; }
		public double MxChar { get; set; }
		public double MxyChar { get; set; }
		public double VxChar { get; set; }
		public double VyChar { get; set; }

		//------------------------------------------------------------------------------------
		//Define EIx EIy, GAx, GAy, ExEquivalent, EyEquivalent to be defined as orthotropic material in Karamba for instance 
		//with the constant stiffness representing the overall thickness of the CLT plate
		//------------------------------------------------------------------------------------



		public CrossSectionCLT(List<int> thicknesses, List<int> orientations, List<IMaterialTimber> materials, int lamellaWidth = 150, bool narrowSideGlued = false)
		{
			LamellaThicknesses = thicknesses;
			LamellaOrientations = orientations;
			LamellaMaterials = materials;
			LamellaWidth = lamellaWidth;
			NarrowSideGlued = narrowSideGlued;
			LamellaDistanceFromTopFibre = new List<double>();
			LamellaDistanceToCDG = new List<double>();
			NbOfLayers = LamellaMaterials.Count;
			ComputeCrossSectionProperties();
		}




		//Implement a picture in Excel / GH equivalent to KLH documentation


		public void ComputeCrossSectionProperties()
		{
			//Checks and preparatory methods
			CheckLayupValidity();
			DefineReferenceElasticityModuls();
			Thickness = LamellaThicknesses.Sum();
			FirstLayer0 = LamellaOrientations.IndexOf(0);
			FirstLayer90 = LamellaOrientations.IndexOf(90);
			LastLayer0 = LamellaOrientations.LastIndexOf(0);
			LastLayer90 = LamellaOrientations.LastIndexOf(90);

			ComputeCenterOfGravities();
			ComputeNetAreas();
			ComputeInertia();


		}

		/// <summary>
		/// Verify that the layup has at least one layer in X and one layer in Y to qualify as CLT and that the geometric
		/// properties can be computed without errors
		/// </summary>
		private void CheckLayupValidity()
		{

			//Check layers orientation
			bool layer0 = false;
			bool layer90 = false;

			for (int i = 0; i < NbOfLayers; i++)
			{
				if (LamellaOrientations[i] == 0) layer0 = true;
				if (LamellaOrientations[i] == 90) layer90 = true;
				if (layer0 && layer90) break;
			}

			if (!layer0 || !layer90) throw new Exception("The CLT layup does not have at least one layer at 0° X and one layer at 90°");

			//Check consistancy in layer input
			if (LamellaMaterials.Count != LamellaOrientations.Count || LamellaMaterials.Count != LamellaThicknesses.Count)
			{
				throw new Exception("The input data is not consistent between the number of Materials, thicknesses and orientations provided");
			}

		}

		private void DefineReferenceElasticityModuls()
		{
			//The reference moduli of elasticity are taken as the first layer in either 0° or 90°
			int i = LamellaOrientations.FindIndex(p => p == 0);
			int j = LamellaOrientations.FindIndex(p => p == 90);
			EcX = LamellaMaterials[i].E0mean;
			EcY = LamellaMaterials[j].E0mean;

		}

		private void ComputeCenterOfGravities()
		{
			double distToCOG = 0;

			double nominatorXX = 0;
			double denominatorXX = 0;
			double nominatorYY = 0;
			double denominatorYY = 0;


			for (int i = FirstLayer0; i < LastLayer0; i++)
			{

			}

			for (int i = FirstLayer90; i < LastLayer90; i++)
			{

			}


			for (int i = 0; i < NbOfLayers; i++)
			{
				double oi = distToCOG + LamellaThicknesses[i] / 2;
				LamellaDistanceFromTopFibre.Add(oi);

				if (LamellaOrientations[i] == 0)
				{
					nominatorXX += LamellaMaterials[i].E0mean / EcX * LamellaThicknesses[i] * oi;
					denominatorXX += LamellaMaterials[i].E0mean / EcX * LamellaThicknesses[i];
				}
				else
				{
					nominatorYY += LamellaMaterials[i].E0mean / EcY * LamellaThicknesses[i] * oi;
					denominatorYY += LamellaMaterials[i].E0mean / EcY * LamellaThicknesses[i];
				}
				distToCOG += LamellaThicknesses[i];

			}

			CenterOfGravityXX = nominatorXX / denominatorXX;
			CenterOfGravityYY = nominatorYY / denominatorYY;

			for (int i = 0; i < NbOfLayers; i++)
			{
				if (LamellaOrientations[i] == 0)
				{
					LamellaDistanceToCDG.Add(LamellaDistanceFromTopFibre[i] - CenterOfGravityXX);
				}
				else
				{
					LamellaDistanceToCDG.Add(LamellaDistanceFromTopFibre[i] - CenterOfGravityYY);
				}
			}
		}

		private void ComputeNetAreas()
		{
			for (int i = 0; i < NbOfLayers; i++)
			{
				if (LamellaOrientations[i] == 0)
				{
					AreaX += LamellaMaterials[i].E0mean / EcX * 1000 * LamellaThicknesses[i];
				}
				else
				{
					AreaY += LamellaMaterials[i].E0mean / EcY * 1000 * LamellaThicknesses[i];
				}

			}
		}

		private void ComputeInertia()
		{
			for (int i = 0; i < NbOfLayers; i++)
			{
				//XX
				if (LamellaOrientations[i] == 0)
				{
					MomentOfInertiaXX += LamellaMaterials[i].E0mean / EcX * 1000 * Math.Pow(LamellaThicknesses[i], 3) / 12;
					MomentOfInertiaXX += LamellaMaterials[i].E0mean / EcX * 1000 * LamellaThicknesses[i] * Math.Pow(LamellaDistanceToCDG[i], 2);
				}

				//YY
				else
				{
					MomentOfInertiaYY += LamellaMaterials[i].E0mean / EcY * 1000 * Math.Pow(LamellaThicknesses[i], 3) / 12;
					MomentOfInertiaYY += LamellaMaterials[i].E0mean / EcY * 1000 * LamellaThicknesses[i] * Math.Pow(LamellaDistanceToCDG[i], 2);
				}
			}


			//Section Modulus
			//XX


			//YY




		}















		public List<double> ComputeNormalStress(double NormalForce)
		{
			throw new NotImplementedException();
		}

		public List<double> ComputeShearStressX(double ShearForceY)
		{
			throw new NotImplementedException();
		}

		public List<double> ComputeShearStressY(double ShearForceZ)
		{
			throw new NotImplementedException();
		}

		public List<double> ComputeStressBendingX(double BendingMomentY)
		{
			throw new NotImplementedException();
		}

		public List<double> ComputeStressBendingY(double BendingMomentZ)
		{
			throw new NotImplementedException();
		}

		public List<double> ComputeTorsionStress(double TorsionMoment)
		{
			throw new NotImplementedException();
		}


	}
}
