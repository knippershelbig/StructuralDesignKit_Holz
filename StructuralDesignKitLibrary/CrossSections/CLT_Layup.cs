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
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public double Thickness { get; set; }
        public int NbOfLayers { get; set; }

        public CrossSectionCLT CS_X { get; set; }
        public CrossSectionCLT CS_Y { get; set; }

        /// <summary>
        /// List of lamella thicknesses
        /// </summary>
        public List<double> Thicknesses { get; set; }

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
        public List<IMaterialTimber> Materials { get; set; }

        /// <summary>
        /// define the with of a single lamella
        /// </summary>
        public int LamellaWidth { get; set; }

        /// <summary>
        /// define if the narrow side of the lamella is glued
        /// </summary>
        public bool NarrowSideGlued { get; set; }


        // Index of the first and last layers at 0° and 90°. Given from the top
        private int FirstLayer0 { get; set; }
        private int FirstLayer90 { get; set; }
        private int LastLayer0 { get; set; }
        private int LastLayer90 { get; set; }


        public CLT_Layup(List<double> thicknesses, List<int> orientations, List<IMaterialTimber> materials, int lamellaWidth = 150, bool narrowSideGlued = false)
        {
            Thicknesses = thicknesses;
            LamellaOrientations = orientations;
            Materials = materials;
            LamellaWidth = lamellaWidth;
            NarrowSideGlued = narrowSideGlued;
            CheckLayupValidity();

            Thickness = Thicknesses.Sum();

            ComputeCrossSections();

        }


        private void ComputeCrossSections()
        {
            FirstLayer0 = LamellaOrientations.IndexOf(0);
            FirstLayer90 = LamellaOrientations.IndexOf(90);
            LastLayer0 = LamellaOrientations.LastIndexOf(0);
            LastLayer90 = LamellaOrientations.LastIndexOf(90);


            List<double> lamellaThicknessX = new List<double>();
            List<int> orientationsX = new List<int>();
            List<IMaterialTimber> materialsX = new List<IMaterialTimber>();

			List<double> lamellaThicknessY = new List<double>();
			List<int> orientationsY = new List<int>();
			List<IMaterialTimber> materialsY = new List<IMaterialTimber>();

			for (int i = FirstLayer0; i < LastLayer0 + 1; i++)
            {
                lamellaThicknessX.Add(Thicknesses[i]);
                orientationsX.Add(LamellaOrientations[i]);
                materialsX.Add(Materials[i]);
            }

            CS_X = new CrossSectionCLT(lamellaThicknessX, orientationsX, materialsX, LamellaWidth, NarrowSideGlued);

            //lamellaThickness.Clear();
            //orientations.Clear();
            //materials.Clear();

            for (int i = FirstLayer90; i < LastLayer90 + 1; i++)
            {
                lamellaThicknessY.Add(Thicknesses[i]);
                if (LamellaOrientations[i] == 0)
                {
                    orientationsY.Add(90);
                }
                else
                {
                    orientationsY.Add(0);
                }
                materialsY.Add(Materials[i]);
            }

            CS_Y = new CrossSectionCLT(lamellaThicknessY, orientationsY, materialsY, LamellaWidth, NarrowSideGlued);
        }

        /// <summary>
        /// Verify that the layup has at least one layer in X and one layer in Y to qualify as CLT and that the geometric
        /// properties can be computed without errors
        /// </summary>
        private void CheckLayupValidity()
        {

            //Check consistancy in layer input
            if (Materials.Count != LamellaOrientations.Count || Materials.Count != Thicknesses.Count)
            {
                throw new Exception("The input data is not consistent between the number of Materials, thicknesses and orientations provided");
            }

            NbOfLayers = Thicknesses.Count;

            //Check layers orientation - Makes sure that the layup is composed of at list one layer in each direction
            bool layer0 = false;
            bool layer90 = false;

            for (int i = 0; i < NbOfLayers; i++)
            {
                if (LamellaOrientations[i] == 0) layer0 = true;
                if (LamellaOrientations[i] == 90) layer90 = true;
                if (layer0 && layer90) break;
            }

            if (!layer0 || !layer90) throw new Exception("The CLT layup does not have at least one layer at 0° X and one layer at 90°");

            if (Thicknesses.Any(p => p <= 0)) throw new Exception("The lamella thicknesses cannot be null or negative");
        }

    }





}
