using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.Vibrations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace StructuralDesignKitExcel.RibbonActions
{
    public static partial class ExcelFormulae
    {

        /// <summary>
        /// Compute the Resonant Response Analysis for Low frequency floors for a given point with multiple vibration modes (RMS acceleration of the response)
        /// </summary>
        /// <param name="List_uen">List of the mode shape amplitudes, from the unity or mass normalised FE output, at the point on the floor where the excitation force Fh is applied</param>
        /// <param name="List_urn">List of the mode shape amplitudes, from the unity or mass normalised FE output, at the point where the response is to be calculated</param>
        /// <param name="NaturalFrequencies">List of natural frequencies for the mode considered</param>
        /// <param name="List_Mg">List of the is the modal mass for the considered modes</param>
        /// <param name="fp">pace frequency in [Hz]</param>
        /// <param name="DampingRatio">Damping as ratio to critical damping</param>
        /// <param name="weigthingCategory">Weighting category for human perception of vibrations</param>
        /// <returns></returns>
        /// 


        [ExcelFunction(Description = "Compute the Resonant Response Analysis for Low frequency floors for a given point with multiple vibration modes (RMS acceleration of the response)",
            Name = "SDK.Vibration.AccelerationResponse",
            IsHidden = false,
            Category = "SDK.Vibration")]
        public static double AccelerationResponse(
            [ExcelArgument(Description = "Range of modes considered")] object[,] Modes,
            [ExcelArgument(Description = "Range of frequencies considered")] object[,] Frequencies,
            [ExcelArgument(Description = "Range of modal masses considered")] object[,] ModalMasses,
            [ExcelArgument(Description = "Data for the standardized displacement according to the documentation")] object[,] StandardizedDisplacements,
            [ExcelArgument(Description = "pace frequency in [Hz]")] double fp,
            [ExcelArgument(Description = "Damping as ratio to critical damping")] double Xi,
            [ExcelArgument(Description = "Weighting category for human perception of vibrations")] string weightingType,
            [ExcelArgument(Description = "Node to consider for the excitation")] int excitationNode,
            [ExcelArgument(Description = "Node to consider for the response")] int ResponseNode,
            [ExcelArgument(Description = "Length of the walking path, if negative, the Eurocode resonant build up factor is considered ")] double walkingLength,
            [ExcelArgument(Description = "If true, provide the Response factor instead of the acceleration")] bool ResponseFactor)
        {
            double accelerationResponse = 0;
            try
            {
                //range to lists
                List<int> modes = new List<int>();
                List<double> frequencies = new List<double>();
                List<double> modalMasses = new List<double>();

                //ensure all lists have the same length
                int listLength = Modes.Length;
                if (Frequencies.Length != listLength || ModalMasses.Length != listLength) throw new ArgumentException("all list should have the same length");

                for (int i = 0; i < Modes.Length; i++)
                {
                    modes.Add(Convert.ToInt32(Modes[i, 0]));
                    frequencies.Add(Convert.ToDouble(Frequencies[i, 0]));
                    modalMasses.Add(Convert.ToDouble(ModalMasses[i, 0]));
                }

                //Divide Data range into 5 lists
                List<int> NodeNb = new List<int>();
                List<int> ModeNb = new List<int>();
                List<double> ux = new List<double>();
                List<double> uy = new List<double>();
                List<double> uz = new List<double>();
                for (int i = 0; i < StandardizedDisplacements.Length / 5; i++)
                {
                    NodeNb.Add(Convert.ToInt32(StandardizedDisplacements[i, 0]));
                    ModeNb.Add(Convert.ToInt32(StandardizedDisplacements[i, 1]));
                    
                    //Currently horizontal response is not implemented
                    //ux.Add(Convert.ToDouble(StandardizedDisplacements[i, 2]));
                    //uy.Add(Convert.ToDouble(StandardizedDisplacements[i, 3]));

                    uz.Add(Convert.ToDouble(StandardizedDisplacements[i, 4]));
                }



                //Create list of displacements to be analized

                int indexNodeExcitation = NodeNb.IndexOf(excitationNode);
                int indexNodeResponse = NodeNb.IndexOf(ResponseNode);

                List<double> uen = new List<double>();
                List<double> urn = new List<double>();

                for (int i = 0; i < listLength; i++)
                {
                    uen.Add(uz[indexNodeExcitation + i]);
                    urn.Add(uz[indexNodeResponse + i]);
                }

                accelerationResponse= Vibrations.ResonantResponseAnalysis(uen, urn, frequencies, modalMasses, fp, Xi, Vibrations.GetWeighting(weightingType),walkingLength, ResponseFactor);


            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            return accelerationResponse;
        }
    }



}

