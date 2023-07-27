using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.CrossSections
{
	/// <summary>
	/// Class representing the composition of a CLT layup
	/// </summary>
	public class CLT_Layup
	{

		/// <summary>
		/// List of lamella thicknesses
		/// </summary>
		public List<int> Thicknesses { get; set; }

		/// <summary>
		/// List of lamella orientation (0° or 90°)
		/// </summary>
		private List<int> _orientations;

		public List<int> Orientations
		{
			get { return _orientations; }
			set
			{
				//Verify if the lamella orientation is properly defined
				foreach (int angle in value)
				{
					if (angle != 0 || angle != 90) throw new Exception("The CLT lamella angle can only be 0° or 90°");
				}
				_orientations = value;
			}
		}

		/// <summary>
		/// List of the material constituing each lamella
		/// </summary>
		public List<IMaterialTimber> Materials { get; set; }

		/// <summary>
		/// define the with of a single lamella
		/// </summary>
		public int LamellaWidth { get; set; }

		/// <summary>
		/// define if the narrow side of the lamella is glued
		/// </summary>
		public bool NarrowSideGlued { get; set; }


		public CLT_Layup(int lamellaWidth = 150, bool narrowSideGlued = false)
		{
			Thicknesses = new List<int>();
			_orientations = new List<int>();
			Materials = new List<IMaterialTimber>();
			LamellaWidth = lamellaWidth;
			NarrowSideGlued = narrowSideGlued;
		}

		public CLT_Layup(List<int> thicknesses, List<int> orientations, List<IMaterialTimber> materials, int lamellaWidth = 150, bool narrowSideGlued = false)
		{
			Thicknesses = thicknesses;
			_orientations = new List<int>();
			Materials = new List<IMaterialTimber>();
			LamellaWidth = lamellaWidth;
			NarrowSideGlued = narrowSideGlued;
		}
	}
}
