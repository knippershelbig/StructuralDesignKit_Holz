
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
    /// <summary>
    /// Class defining a CLT Cross-section
    /// A cross section differs from a layup. The layup represents the "physical" assembly of boards
    /// while whe cross section provides the mechanical properties, either in X (0°) or Y (90°)
    /// 
    /// the cross section can be equivalent to its parent layup or it can be a part of it
    /// 
    /// A layup two cross sections 
    /// The different notations and approaches are taken from 
    /// - "2018 Wallner-Novak M., CLT structural design I proHOLZ -  ISBN 978-3-902926-03-6"
    /// - "2018 Wallner-Novak M., CLT structural design II proHOLZ -  ISBN 978-3-902320-96-4"
    /// </summary>
    public class CrossSectionCLT
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public double Thickness { get; set; }
        public int NbOfLayers { get; set; }

        /// <summary>
        /// List of lamella thicknesses
        /// </summary>
        public List<double> LamellaThicknesses { get; set; }

        /// <summary>
        /// List of lamella orientation (0° or 90°)
        /// </summary>

        public List<int> LamellaOrientations { get; set; }

        /// <summary>
        /// List of the material constituing each lamella
        /// </summary>
        public List<IMaterialTimber> LamellaMaterials { get; set; }

        /// <summary>
        /// Position of the lamella center of gravity from the CS upper edge
        /// </summary>
        public List<double> LamellaCoGDistanceFromTopFibre { get; set; }

        /// <summary>
        /// Distance from the center of gravity of a layer toward the overall center of gravity
        /// </summary>
        public List<double> LamellaDistanceToCDG { get; set; }


        //Reference E modulus 
        private double ERef { get; set; }



        //Center of gravity in both main direction ; in mm from the top
        public double CenterOfGravity { get; set; }


        /// <summary>
        /// Distance of the bottom edge to the overall center of gravity - X direction
        /// </summary>
        public double Zu { get; set; }

        /// <summary>
        /// Distance of the top edge to the overall center of gravity - X direction
        /// </summary>
        public double Zo { get; set; }

        /// <summary>
        /// Active area in the considered direction in mm²
        /// </summary>
        public double Area { get; set; }

        /// <summary>
        /// Moment of inertia in mm4
        /// </summary>
        public double MomentOfInertia { get; set; }

        /// <summary>
        /// Section Module (net) in mm³
        /// </summary>
        public double W0_Net { get; set; }

        /// <summary>
        /// Static Moment for rolling shear
        /// </summary>
        public double Sr0Net { get; set; }


        /// <summary>
        /// Static Moment for  shear
        /// </summary>
        public double S0Net { get; set; }

        public double TorsionalInertia { get; set; }
        public double TorsionalModulus { get; set; }
        public double NCompressionChar { get; set; }
        public double NTensionChar { get; set; }
        public double NxyChar { get; set; }
        public double MChar { get; set; }
        public double MxyChar { get; set; }
        public double VChar { get; set; }



        //------------------------------------------------------------------------------------
        //Define EIx EIy, GAx, GAy, ExEquivalent, EyEquivalent to be defined as orthotropic material in Karamba for instance 
        //with the constant stiffness representing the overall thickness of the CLT plate
        //------------------------------------------------------------------------------------



        public CrossSectionCLT(List<double> thicknesses, List<int> orientations, List<IMaterialTimber> materials, int lamellaWidth = 150, bool narrowSideGlued = false)
        {
            LamellaThicknesses = thicknesses;
            LamellaOrientations = orientations;
            LamellaMaterials = materials;
            LamellaCoGDistanceFromTopFibre = new List<double>();
            LamellaDistanceToCDG = new List<double>();
            NbOfLayers = LamellaMaterials.Count;
            ComputeCrossSectionProperties();
        }


        //Implement a picture in Excel / GH equivalent to KLH documentation


        public void ComputeCrossSectionProperties()
        {
            //Define the top lamella as the reference material
            ERef = LamellaMaterials[0].E0mean;
            Thickness = LamellaThicknesses.Sum();

            CenterOfGravity = ComputeCenterOfGravity(this);
            ComputeNetAreas();
            ComputeInertia();
            ComputeSectionModulus();
            ComputeStaticMoment();
        }



        private double ComputeCenterOfGravity( CrossSectionCLT CS)
        {
            double distToCOG = 0;

            double nominator = 0;
            double denominator = 0;

            for (int i = 0; i < CS.NbOfLayers; i++)
            {
                //Position of the center of gravity of the individual layers from the element's upper edge
                double oi = distToCOG + CS.LamellaThicknesses[i] / 2;

                CS.LamellaCoGDistanceFromTopFibre.Add(oi);

                if (CS.LamellaOrientations[i] == 0)
                {
                    nominator += CS.LamellaMaterials[i].E0mean / CS.ERef * CS.LamellaThicknesses[i] * oi;
                    denominator += CS.LamellaMaterials[i].E0mean / CS.ERef * CS.LamellaThicknesses[i];
                }

                distToCOG += CS.LamellaThicknesses[i];
            }

             CS.CenterOfGravity =  nominator / denominator;


            //Compute the distance of the center of gravity of the individual layers from the Layup CoG
            for (int i = 0; i < CS.NbOfLayers; i++)
            {
                CS.LamellaDistanceToCDG.Add(CS.LamellaCoGDistanceFromTopFibre[i] - CS.CenterOfGravity);
            }


            //Compute the distance of the top and bottom edge toward the center of gravity
            CS.Zo = Math.Abs(CS.LamellaDistanceToCDG.First()) + CS.LamellaThicknesses.First() / 2;
            CS.Zu = Math.Abs(CS.LamellaDistanceToCDG.Last()) + CS.LamellaThicknesses.Last() / 2;

            return CS.CenterOfGravity;
        }

        /// <summary>
        /// Compute the net area in both directions
        /// </summary>
        private void ComputeNetAreas()
        {
            for (int i = 0; i < NbOfLayers; i++)
            {
                if (LamellaOrientations[i] == 0)
                {
                    Area += LamellaMaterials[i].E0mean / ERef * 1000 * LamellaThicknesses[i];
                }
            }
        }


        private void ComputeInertia()
        {
            for (int i = 0; i < NbOfLayers; i++)
            {
                if (LamellaOrientations[i] == 0)
                {
                    MomentOfInertia += LamellaMaterials[i].E0mean / ERef * 1000 * Math.Pow(LamellaThicknesses[i], 3) / 12;
                    MomentOfInertia += LamellaMaterials[i].E0mean / ERef * 1000 * LamellaThicknesses[i] * Math.Pow(LamellaDistanceToCDG[i], 2);
                }
            }
        }



        private void ComputeStaticMoment()
        {
            //define longitudinal layer closest to the position of the center of gravity
            int lamellaIndex = 0;
            bool CoGinLayer = false;
            for (int i = 0; i < LamellaDistanceToCDG.Count; i++)
            {
                if (LamellaOrientations[i] == 0)
                {
                    if (Math.Abs(LamellaDistanceToCDG[i]) < Math.Abs(LamellaDistanceToCDG[lamellaIndex])) lamellaIndex = i;
                }
            }


            for (int i = 0; i < lamellaIndex + 1; i++)
            {
                if (LamellaOrientations[i] == 0)
                {
                    Sr0Net += ERef / LamellaMaterials[i].E0mean * 1000 * LamellaThicknesses[i] * Math.Abs(LamellaDistanceToCDG[i]);
                }
            }

            //define if Center of gravity is situated in the closest longitudinal layer 
            if (LamellaDistanceToCDG[lamellaIndex] < LamellaThicknesses[lamellaIndex] / 2) CoGinLayer = true;

            //if the center of gravity is located in the affected region
            if (CoGinLayer)
            {
                for (int i = 0; i < lamellaIndex + 1; i++)
                {
                    if (LamellaOrientations[i] == 0)
                    {
                        S0Net += ERef / LamellaMaterials[i].E0mean * 1000 * LamellaThicknesses[i] * LamellaDistanceToCDG[i] + 1000 * (Math.Pow(LamellaThicknesses[lamellaIndex] / 2 - LamellaDistanceToCDG[lamellaIndex], 2) / 2);
                    }
                }
          
            }

            //if the center of gravity is not situated in the closest longitudinal layer
            else
            {
                S0Net = Sr0Net;
            }


        }


        //Compute the different static moments as follow
        //3 values per lamella up (top / middle / bottom) unless the CoG is in the lamella. Then a 4rth value is added
        //The calculation is done from top to CoG and then from Bottom to COG
        //The values are then sorted from top to bottom, per lamella
        private List<List<double>> ComputeStaticMomentPerLamella()
        {

            return new List<List<double>>();
        }


        private void ComputeSectionModulus()
        {
            W0_Net = MomentOfInertia / Math.Max(Math.Abs(Zu), Math.Abs(Zo));
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
