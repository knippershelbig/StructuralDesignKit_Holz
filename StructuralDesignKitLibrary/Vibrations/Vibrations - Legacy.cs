using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace StructuralDesignKitLibrary.Vibrations
{
    public static class VibrationsLegacy
    {


        //below 8Hz : acceleration criteria - arms < Response factor × arms,base when f1 < 8 Hz 
        //Above 8Hz : Velocity criteria - vrms < Response factor × vrms,base when f1 ≥ 8 Hz 
        public static double Acc()
        {


            //walking frequency min FwMin
            //walking frequency max FwMax
            //Static weight of the walker P
            //Stride length of the walker L
            //Span length of the bay in which excitation node reside
            //Damping as ratio to critical damping

            //---------------------------------------------------
            //Resonant Response Analysis(Low - Frequency Floors)
            //---------------------------------------------------

            //For resonant response analysis, CCIP-016 recommends considering all modes with modal frequencies up to 15 Hz

            double step = 0.01; //step in Hz for the analysis; Should not be greater than 0.025Hz. The lower the more precise will be the analysis (but also the longuer)

            int h = 1;              //Harmonic number; an integer not less than 1 used to fnd harmonic frequencies of loading frequency, unitless
            double fw = 2;        //Walking frequency, Hz
            double fh = fw * h;     //Harmonic frequency of loading, harmonic number times walking frequency, h*fw, Hz
            int P = 760;             //Weight of Walker in Kg -  AISC Design Guide 11 - EC5 considers 70Kg - SCI Considers 76
            double l = 0.75;        //Stride length - usually between 0.6 to 0.9m



            //Harmonic force for resonant analysis
            //This force represents the amplitude of the forcing function derived empirically from tests of walkers on instrumented platforms.
            //Fh is the dynamic component of the total force applied by the walker to the foor (i.e., it is the total load minus the static weight of the walker

            double Fh = HarmonicCoefficient(fh, h) * P;

            double Xi = 0.03; //modal damping ratio
            double L = 5.2; //Floor span in m

            double Nh = 0.55 * h * L / l; //alculated number of loading cycles, e.g., steps

            //Sub-resonant response reduction factor for limited walking distance for a harmonic and mode, unitless
            //Depending on the walking path and space planning, it is possible that walkers have crossed and exited the space before the appropriate number of steps
            //(i.e., loading cycles) have taken place to achieve steady-state response.CCIP - 016 proposes a subresonant correction factor, ρh, m, to account for this effect;
            //however, it is not usually found to be infuential in foors and can be conservatively taken as unity.
            double Rhohm = 1 - Math.Exp(-2 * Math.PI * Xi * Nh);


            //for a given mode and harmonic the resonant response accelerations:
            double fm = 18.254; //the natural frequency of the mode under consideration (up to 15 Hz)


            //Per node, for each mode and for each harmonic:

            double urm = 1;
            double uem = 1;
            double mhat = 5000;

            double Am = 1 - Math.Pow(fh / fm, 2);
            double Bm = 2 * Xi * fh / fm;

            double a_real_h_m = Math.Pow(fh / fm, 2) * Fh * urm * uem * Rhohm / mhat * Am / (Math.Pow(Am, 2) * Math.Pow(Bm, 2));
            double a_imag_h_m = Math.Pow(fh / fm, 2) * Fh * urm * uem * Rhohm / mhat * Bm / (Math.Pow(Am, 2) * Math.Pow(Bm, 2));

            double ah = Math.Sqrt(Math.Pow(a_real_h_m, 2) + Math.Pow(a_imag_h_m, 2));


            return ah;



        }


        


        /// <summary>
        /// Impulsive excitation of walking according to SCI P354 - Eq 18
        /// </summary>
        /// <param name="fp">Walking frequency</param>
        /// <param name="fn">Floor modal frequency</param>
        /// <returns></returns>
        public static double MeanNodalImpulseSCI(double fp, double fn)
        {
            return 60 * 746 / 700 * Math.Pow(fp, 1.43) / Math.Pow(fn, 1.3);
        }


        /// <summary>
        /// Mean modal impulse - footfall idealised as a mean impulsive load in Ns according to Draft prEN 1995-1-1 - §9.3.7 - Eq9.12
        /// </summary>
        /// <param name="fw">Walking frequency</param>
        /// <param name="fn">Floor modal frequency</param>
        /// <returns></returns>
        public static double MeanNodalImpulseEC5(double fw, double fn)
        {
            return 42 * Math.Pow(fw, 1.43) / Math.Pow(fn, 1.3);
        }


        public static double HarmonicCoefficient(double fh, int HarmonicNumber)
        {
            double harmonicCoefficient = 0;
            if (HarmonicNumber <= 0) throw new Exception("HarmonicNumber cannot be negative or null");
            else if (HarmonicNumber == 1) harmonicCoefficient = Math.Min(0.56, 0.41 * (fh - 0.95));
            else if (HarmonicNumber == 2) harmonicCoefficient = 0.069 + 0.0056 * fh;
            else if (HarmonicNumber == 3) harmonicCoefficient = 0.033 + 0.0064 * fh;
            else if (HarmonicNumber == 4) harmonicCoefficient = 0.013 + 0.0065 * fh;
            else if (HarmonicNumber > 4) harmonicCoefficient = 0;

            return harmonicCoefficient;
        }
    }
}
