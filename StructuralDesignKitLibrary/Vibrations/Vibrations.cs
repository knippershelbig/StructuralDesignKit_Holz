using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;
using static StructuralDesignKitLibrary.Vibrations.Vibrations;


namespace StructuralDesignKitLibrary.Vibrations
{
    public static class Vibrations
    {
        //below 8Hz : acceleration criteria - arms < Response factor × arms,base when f1 < 8 Hz 

        //Above 8Hz : Velocity criteria - vrms < Response factor × vrms,base when f1 ≥ 8 Hz 


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
        /// <param name="walkingLength">Length of the walking path, if negative, the Eurocode resonant build up factor is considered</param>
        /// <param name="ResponseFactor">If true, provide the Response factor instead of the acceleration</param>
        /// <returns></returns>
        public static double ResonantResponseAnalysis(List<double> List_uen, List<double> List_urn, List<double> NaturalFrequencies, List<double> List_Mg, double fp, double DampingRatio, Weighting weigthingCategory, double walkingLength = -1, bool ResponseFactor = false)
        {
            int Q = 746; //Static force exerted by an average person normally taken as 76kg x 9.81m/s² = 746 N
            double Xi = DampingRatio;

            double BuildupFactor = ComputeResonanceBuildupFactor(walkingLength, fp, DampingRatio);

            List<List<double>> accelerationResponse = new List<List<double>>();

            //Iterate over 4 harmonics
            for (int i = 1; i < 5; i++)
            {
                int h = i;

                double fh = fp * h;                             //Harmonic frequency of loading, harmonic number times walking frequency, h*fw, Hz
                double Fh = HarmonicCoefficient(fh, h) * Q;     //Excitation force for the harmonic considered

                List<double> responseMode = new List<double>();
                int modeCount = 0;

                //iterate over each mode
                foreach (double fn in NaturalFrequencies)
                {
                    double W = ComputeWeightingFactor(weigthingCategory, fn);
                    responseMode.Add(ComputeAccelerationResponse(List_uen[modeCount], List_urn[modeCount], Fh, List_Mg[modeCount], ComputeDnh(h, Xi, fp, fn), W) * BuildupFactor);
                    modeCount++;
                }
                accelerationResponse.Add(responseMode);
            }

            if (ResponseFactor) return SRSSAcceleration(accelerationResponse) / 0.005;
            else return SRSSAcceleration(accelerationResponse);
        }


        public static double TransientResponseAnalysis(List<double> List_uen, List<double> List_urn, List<double> NaturalFrequencies, List<double> List_Mg, double fp, double DampingRatio, Weighting weigthingCategory, bool ResponseFactor, double resolution = 0.00125)
        {

            
            int steps = (Int32)Math.Round(1 / fp / resolution);
            double timeStep = 1 / fp / steps;


            List<List<double>> VelocitySeriesPerMode = new List<List<double>>();
            int modeCount = 0;


            foreach (double modeFrequence in NaturalFrequencies)
            {
                VelocitySeriesPerMode.Add(ComputeVelocityResponseTimeSeries(List_uen[modeCount], List_urn[modeCount], List_Mg[modeCount], modeFrequence, fp, DampingRatio, weigthingCategory, resolution));
                modeCount++;
            }

            double transientResponse = Vrms(VelocitySeriesPerMode, timeStep, fp);
            if (ResponseFactor) transientResponse /= 0.0001;
            return transientResponse;
        }


        public static double ComputeResonanceBuildupFactor(double walkingLength, double fp, double Xi)
        {
            double resonanceBuildupFactor = 0;
            if (walkingLength > 0)
            {
                if (fp >= 1.7 && fp <= 2.4)
                {
                    //Walking speed according to Bachmann and Ammann for walking pace between 1.7Hz and 2.4Hz; SCI P354 - Eq 16
                    double v = 1.67 * Math.Pow(fp, 2) - 4.83 * fp + 4.5;
                    resonanceBuildupFactor = 1 - Math.Exp(-2 * Math.PI * Xi * walkingLength / v);
                }
            }
            //Outside the boundaries defined by the equation or if walking length defined as negative,
            //the normalized value of 0.4 according to the Pr EN 1995-1-1 §9.3.6 is chosen
            else resonanceBuildupFactor = 0.4;

            return resonanceBuildupFactor;
        }


        public static List<double> ComputeVelocityResponseTimeSeries(double uen, double urn, double Mn, double fn, double fp, double Xi, Weighting weighting, double resolution = 0.00125)
        {
            double steps = (Int32)Math.Round(1 / fp / resolution);
            double timeStep = 1 / fp / steps;

            List<double> velocitySeries = new List<double>();

            double W = ComputeWeightingFactor(weighting, fn);
            int Q = 746;             //Static force exerted by an average person normally taken as 76kg x 9.81m/s² = 764 N
            double Fi = MeanNodalImpulseEC5(fp, fn);//Equivalent impulsive force representing a single footfall in Ns


            double t = 0;
            for (int i = 0; i < steps; i++)
            {
                velocitySeries.Add(VelocityResponse(uen, urn, fn, fp, Fi, W, Xi, t, Mn, weighting));
                t += timeStep;
            }

            return velocitySeries;
        }


        public static double VelocityResponse(double uen, double urn, double fn, double fp, double Fi, double W, double Xi, double t, double Mn, Weighting weighting)
        {



            return uen * urn * Fi / Mn * Math.Sin(2 * Math.PI * fn * Math.Sqrt(1 - Math.Pow(Xi, 2)) * t) * Math.Exp(-2 * Math.PI * Xi * fn * t) * W;
        }


        /// <summary>
        /// Root mean square of the velocity
        /// </summary>
        /// <returns></returns>
        public static double Vrms(List<List<double>> velocityResponses, double timeStep, double fp)
        {

            //Check if lists have the same length
            int initLength = velocityResponses[0].Count;
            foreach (List<double> list in velocityResponses)
            {
                if (list.Count != initLength) throw new Exception("All list must have the same length");
            }

            //Velocity response from impulse loading as a function time for mode:
            List<double> Vwer_t = new List<double>();
            for (int i = 0; i < initLength; i++)
            {
                double sum = 0;
                for (int j = 0; j < velocityResponses.Count; j++)
                {
                    sum += velocityResponses[j][i];
                }
                Vwer_t.Add(sum);
            }
            List<double> VrmsList = new List<double>();
            foreach (double v in Vwer_t)
            {
                VrmsList.Add(Math.Pow(v, 2) * timeStep);
            }
            double Vrms = VrmsList.Sum();
            Vrms = Math.Sqrt(Vrms * fp);
            return Vrms;
        }


        /// <summary>
        /// Impulsive excitation of walking according to SCI P354 - Eq 18
        /// </summary>
        /// <param name="fp">Walking frequency</param>
        /// <param name="fn">Floor modal frequency</param>
        /// <returns></returns>
        private static double MeanNodalImpulseSCI(double fp, double fn, double Q)
        {
            return 60 * Q / 700 * Math.Pow(fp, 1.43) / Math.Pow(fn, 1.3);
        }


        /// <summary>
        /// Mean modal impulse - footfall idealised as a mean impulsive load in Ns according to Draft prEN 1995-1-1 - §9.3.7 - Eq9.12
        /// </summary>
        /// <param name="fp">Walking frequency</param>
        /// <param name="fn">Floor modal frequency</param>
        /// <returns></returns>
        private static double MeanNodalImpulseEC5(double fp, double fn)
        {
            return 42 * Math.Pow(fp, 1.43) / Math.Pow(fn, 1.3);
        }

        /// <summary>
        /// design Fourier coefficients for walking activities according to SCI P354 Table 3.1
        /// </summary>
        /// <param name="fh">Walking frequency</param>
        /// <param name="HarmonicNumber"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static double HarmonicCoefficient(double fh, int HarmonicNumber)
        {
            double harmonicCoefficient = 0;
            if (HarmonicNumber <= 0 || HarmonicNumber > 4) throw new Exception("HarmonicNumber should be betwwen 1 and 4");
            else if (HarmonicNumber == 1) harmonicCoefficient = 0.436 * (fh - 0.95);
            else if (HarmonicNumber == 2) harmonicCoefficient = 0.006 * (fh + 12.3);
            else if (HarmonicNumber == 3) harmonicCoefficient = 0.007 * (fh + 5.2);
            else if (HarmonicNumber == 4) harmonicCoefficient = 0.007 * (fh + 2);


            return harmonicCoefficient;
        }


        /// <summary>
        /// Dynamic mignification factor for acceleration; Ratio of the peak amplitude to the static amplitude
        /// </summary>
        /// <param name="h">harmonic mode from 1 to 4</param>
        /// <param name="Xi">damping ratio</param>
        /// <param name="fp">pace frequency, Hz</param>
        /// <param name="fn">Floor modal frequency</param>
        /// <returns></returns>
        private static double ComputeDnh(int h, double Xi, double fp, double fn)
        {
            double BetaN = fp / fn;
            return Math.Pow(h, 2) * Math.Pow(BetaN, 2) / Math.Sqrt(Math.Pow(1 - Math.Pow(h, 2) * Math.Pow(BetaN, 2), 2) + Math.Pow(2 * h * Xi * BetaN, 2));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="uen">mode shape amplitudes, from the unity or mass normalised FE output, at the point on the floor where the excitation force Fh is applied</param>
        /// <param name="urn">mode shape amplitudes, from the unity or mass normalised FE output, at the point where the response is to be calculated</param>
        /// <param name="Fh">Excitation force for the harmonic considered</param>
        /// <param name="Mn">modal mass for the considered mode</param>
        /// <param name="Dnh">Dynamic mignification factor for acceleration; Ratio of the peak amplitude to the static amplitude</param>
        /// <param name="W">Code-defined weighting factor for human perception of vibrations</param>
        /// <returns></returns>
        private static double ComputeAccelerationResponse(double uen, double urn, double Fh, double Mn, double Dnh, double W)
        {
            return uen * urn * Fh / Mn * Dnh * W;
        }


        /// <summary>
        /// Compute the Square-root sum of squares (SRSS) to sum the accelerations of a considered point
        /// </summary>
        /// <param name="aerh"></param>
        /// <returns></returns>
        private static double SRSSAcceleration(List<List<double>> aerh)
        {
            double aw_rms_e_r = 0;
            List<double> AccelerationResponsesHarmonic = new List<double>();

            foreach (List<double> modes in aerh)
            {
                double acceleration = 0;
                foreach (double acc_Mode in modes)
                {
                    acceleration += acc_Mode;
                }
                AccelerationResponsesHarmonic.Add(Math.Pow(acceleration, 2));
            }
            foreach (double acc in AccelerationResponsesHarmonic)
            {
                aw_rms_e_r += acc;
            }

            return Math.Sqrt(aw_rms_e_r) / Math.Sqrt(2);

        }


        /// <summary>
        /// Weighting Categories according to SCI P354 - Table 5.1
        /// </summary>
        public enum Weighting
        {
            Workshop,
            CirculationSpace,
            Residential,
            Office,
            Ward,
            GeneralLaboratory,
            ConsultingRoom,
            CriticalWorkingArea,
            None
        }

        /// <summary>
        /// Helper function to get a Weighting Enum based on a string
        /// </summary>
        /// <param name="weighting">Service Class</param>
        /// <returns></returns>
        public static Weighting GetWeighting(string weighting)
        {
            switch (weighting)
            {
                case "Workshop":
                    return Weighting.Workshop;
                case "CirculationSpace":
                    return Weighting.CirculationSpace; ;
                case "Residential":
                    return Weighting.Residential; ;
                case "Office":
                    return Weighting.Office; ;
                case "Ward":
                    return Weighting.Ward; ;
                case "GeneralLaboratory":
                    return Weighting.GeneralLaboratory; ;
                case "ConsultingRoom":
                    return Weighting.ConsultingRoom; ;
                case "CriticalWorkingArea":
                    return Weighting.CriticalWorkingArea; ;
                case "None":
                    return Weighting.None; ;
                default:
                    return Weighting.None; ;
            }
        }

        /// <summary>
        /// Compute the weighting factors appropriate for floor design according to SCI P354 - Table 5.1
        /// </summary>
        /// <param name="weighting"></param>
        /// <param name="f"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static double ComputeWeightingFactor(Weighting weighting, double f)
        {
            double W = 1;
            //Weighting Wg according to SCI P354 - Table 5.1
            if (weighting == Weighting.CirculationSpace || weighting == Weighting.Workshop)
            {
                if (f > 1 && f < 4) W = 0.5 * Math.Sqrt(f);
                else if (f >= 4 && f <= 8) W = 1;
                else if (f > 8) W = 8 / f;
                else throw new Exception("frequency not covered in weighting Wg");
            }
            else if (weighting == Weighting.Residential || weighting == Weighting.Office || weighting == Weighting.Ward || weighting == Weighting.GeneralLaboratory || weighting == Weighting.ConsultingRoom)
            {
                if (f > 1 && f < 2) W = 0.4;
                else if (f >= 2 && f < 5) W = f / 5;
                else if (f >= 5 && f <= 16) W = 1;
                else if (f > 16) W = 16 / f;
                else throw new Exception("frequency not covered in weighting Wb");
            }
            else if (weighting == Weighting.CriticalWorkingArea)
            {
                if (f > 1 && f < 2) W = 1;
                else if (f >= 2) W = 2 / f;
                else throw new Exception("frequency not covered in weighting Wb");
            }

            else if (weighting == Weighting.CriticalWorkingArea) W = 1;

            return W;
        }
    }
}
