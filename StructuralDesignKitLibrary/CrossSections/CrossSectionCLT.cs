using Dlubal.RFEM5;
using StructuralDesignKitLibrary.CrossSections.Interfaces;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
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
		public CLT_Layup Layup { get; set; }
		public double AreaX { get; set; }
		public double AreaY { get; set; }
		public double MomentOfInertia_XX { get; set; }
		public double MomentOfInertia_YY { get; set; }
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


		//Implement a picture in Excel / GH equivalent to KHL documentation



		public void ComputeCrossSectionProperties()
		{
			throw new NotImplementedException();
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
