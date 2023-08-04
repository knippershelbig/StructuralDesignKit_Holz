using Dlubal.WS.Rfem6.Model;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.Utilities;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.EC5
{
    public static class EC5_Utilities
    {

        public enum TimberType
        {
            Softwood,
            Hardwood,
            Glulam,
            LVL,
            Baubuche,
        }

        public enum ServiceClass
        {
            SC1,
            SC2,
            SC3
        }

        public enum LoadDuration
        {
            Permanent,
            LongTerm,
            MediumTerm,
            ShortTerm,
            Instantaneous,
            ShortTerm_Instantaneous,
        }

        public enum FastenerType
        {
            Bolt,
            Dowel
            //Nail,
            //Staple,
            //Screw,
            //Bulldog,
            //Ring,
            //...
        }

        #region FireDesign
        /// <summary>
        /// Return the effective charring depth of an exposed timber linear member (beam, column, ...). Result is in [mm] per exposed face
        /// </summary>
        /// <param name="t">exposed duration</param>
        /// <param name="timber">Timber member material</param>
        /// <returns></returns>
        public static double ComputeCharringDepthUnprotectedBeam(int t, IMaterialTimber timber)
        {
            //No strength layer according to DIN EN 1995-1-2 §4.2.2
            double d0 = 7.0;

            double d_char = timber.Bn * t;
            double k0 = ComputeK0(0, t, false);
            double d_eff = d_char + k0 * d0;

            return d_eff;
        }

        //Charring calculation for Panels
        public static double ComputeCharringDepthUnprotectedPanel(int t, IMaterialTimber timber)
        {
            throw new Exception("Method not yet Implemented");
        }

        /// <summary>
        /// Return the effective charring depth of a protected timber linear member (beam, column, ...). Result is in [mm] per exposed face
        /// </summary>
        /// <param name="t">exposed duration</param>
        /// <param name="timber">Timber member material</param>
        /// <param name="plasterboards">List of the type of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="plasterboardthicknesses">List of the thicknesses of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="closedJoint">if true, joints between plasterboards is considered <2mm; if false > 2mm</param>
        /// <param name="horizontal">Boolean value to represent if the protection is horizontal or vertical</param>
        /// <returns></returns>
        public static double ComputeCharringDepthProtectedBeam(int t, IMaterialTimber timber, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, bool closedJoint, bool horizontal)
        {
            if (plasterboardthicknesses.Count > 2 || plasterboards.Count > 2) throw new Exception("Standardized protection in the EN 1995-1-2 §3.4.3.3 accounts only for a maximum of 2 boards");

            double d0 = 7.0;

            double tch = ComputeCombustionStart(plasterboards, plasterboardthicknesses, closedJoint);

            double tf = ComputePanelFailureTime(tch, plasterboards, plasterboardthicknesses, horizontal);

            double K2 = ComputeK2(plasterboards, plasterboardthicknesses);

            double Bn = timber.Bn;

            double ta = ComputeTa(tch, tf, K2, Bn);

            double k0 = ComputeK0(tch, t, true);


            //Due to the implementation of different sources for the failure time of the protection (ÖNORM, Draft EN-1995-2 [2020])
            //The calculation of the protection failure time can be shorter than the combustion start.
            //In this case, we set tch = tf
            if (tf < tch) tch = tf;

            // t < tch
            if (t < tch) return 0;

            // tch < t < ta
            else if (t >= tch && t < tf) return K2 * Bn * (t - tch) + k0 * d0;

            //tf < t < ta
            else if (t >= tf && t < ta) return K2 * Bn * (tf - tch) + K3 * Bn * (t - tf) + k0 * d0;

            // ta < t
            //else return Bn * (t - ta) + Math.Min(25, Bn * ta) + k0 * d0;
            else return Bn * (t - ta) + K2 * Bn * (tf - tch) + K3 * Bn * (ta - tf) + k0 * d0;
        }


        //Reduction factor based on the beginning of combustion
        public static double ComputeK0(double tch, int t, bool protectedSurface)
        {
            double k0;

            if (protectedSurface && t > 20)
            {
                k0 = Math.Min(SDKUtilities.LinearInterpolation(t, 0, 0, tch, 1), 1);
            }
            else
            {
                if (t < 20) k0 = (double)t / 20;
                else k0 = 1;
            }
            return k0;
        }

        /// <summary>
        /// Start of combustion of protected surface according to EN 1995-1-2 §3.4.3.3
        /// </summary>
        /// <param name="plasterboards">List of the type of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="plasterboardthicknesses">List of the thicknesses of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="closedJoint">if true, joints between plasterboards is considered <2mm; if false > 2mm</param>
        /// <returns></returns>
        public static double ComputeCombustionStart(List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, bool closedJoint)
        {
            if (plasterboardthicknesses.Count > 2 || plasterboards.Count > 2) throw new Exception("Standardized protection in the EN 1995-1-2 §3.4.3.3 accounts only for a maximum of 2 boards");

            double tch = 0;
            double hp = 0; // Effective thickness

            //Define effective thickness
            //Case single layer
            if (plasterboardthicknesses.Count == 1) hp = plasterboardthicknesses[0];

            //Case 2 layers
            else
            {
                // case type A or H or A/H + F (not covered by EN 1995-1-2 §3.4.3.3 -> Same approach as A/H for conservative calculation)
                if (plasterboards[0] != plasterboards[1] || plasterboards[0] == PlasterboardType.TypeA) hp = plasterboardthicknesses[0] + plasterboardthicknesses[1] * 0.5;

                // case type F    
                else hp = plasterboardthicknesses[0] + plasterboardthicknesses[1] * 0.8;
            }


            if (closedJoint)
            {
                tch = 2.8 * hp - 14; //DIN EN 1995-1-2 Eq 3.11
            }
            else
            {
                tch = 2.8 * hp - 23; //DIN EN 1995-1-2 Eq 3.12

            }

            return tch;
        }


        /// <summary>
        /// Compute the Panel Failure Time according to both the ÖNORM and the Draft EN 1995-1-2:2020 (E) to provide a safer design while 
        /// waiting for the implementation of the new version of the Eurocode 5. Variations from the ÖNORM apply for horizontal plasterboards type A/H
        /// </summary>
        /// <param name="tch"></param>
        /// <param name="plasterboards"></param>
        /// <param name="plasterboardthicknesses"></param>
        /// <param name="horizontal">Boolean value to represent if the protection is horizontal or vertical</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static double ComputePanelFailureTime(double tch, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, bool horizontal)
        {
            //--------------------------------------------------------------------------------------------------------------------------------
            //NOTES ON THE CALCULATION OF THE PANEL FAILURE TIME
            //--------------------------------------------------------------------------------------------------------------------------------
            //As the current version of the Eurocode 5 does not provide a mean to calculate tf, the Austrian approach based on the 
            //results from the Holzforschung Austria[Teibinger und Matzinger 2010] is implemented.
            //However, the ÖNORM calculation provides inconsistent results on horizontal surfaces where the plasterboards type A perform better than
            //the plasterboards type F.
            //To take this shortcoming into account, the panel failure time according to the draft of the new EN 1995-1-2:2020 (E) is considered for
            //the calculation of horizontal plasterboard type A/H
            //--------------------------------------------------------------------------------------------------------------------------------

            //Sanity checks
            if (plasterboards.Count > 2 || plasterboards.Count <= 0) throw new Exception("The amount of plasterboard should be 1 or 2");
            if (!plasterboards.Contains(PlasterboardType.TypeF) && !plasterboards.Contains(PlasterboardType.TypeA)) throw new Exception("The amount of plasterboard should be 1 or 2");


            //tf for plasterboards type A/H
            if (plasterboards.Contains(PlasterboardType.TypeA))
            {
                if (plasterboards.Count == 1 && horizontal)
                {
                    double hp = plasterboardthicknesses.Sum();

                    return (2.1 * hp - 9) * 1.1; //EN 1995-1-2:2020 (E) Table 5.4
                }

                else if (plasterboards.Count == 2 && horizontal)
                {
                    double hp = plasterboardthicknesses.Sum();

                    return (1.7 * hp - 13) * 1.1; //EN 1995-1-2:2020 (E) Table 5.4
                }

                else return tch; //EN 1995-1-2 §3.4.3.4 (3)
            }

            //tf for plasterboards type F according to  Holzforschung Austria[Teibinger und Matzinger 2010]
            else
            {
                double hp;
                double tf_p = 0;

                if (plasterboards.Count == 1)
                {
                    hp = plasterboardthicknesses[0];
                }
                else
                {
                    hp = plasterboardthicknesses[0] + plasterboardthicknesses[1] * 0.8;
                }

                if (horizontal) tf_p = 1.4 * hp + 6;

                else tf_p = 2.2 * hp + 4;

                return tf_p;
            }
        }



        /// <summary>
        /// Return the minimum fire protection fastener length in [mm] according to EN 1995-1-2 §3.4.3.4 (4) - Eq 3.16
        /// </summary>
        /// <param name="d">Fastener diameter in [mm]</param>
        /// <param name="t">exposed duration</param>
        /// <param name="timber">Timber member material</param>
        /// <param name="plasterboards">List of the type of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="plasterboardthicknesses">List of the thicknesses of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="closedJoint">if true, joints between plasterboards is considered <2mm; if false > 2mm</param>
        /// <param name="horizontal">Boolean value to represent if the protection is horizontal or vertical</param>
        /// <returns></returns>
        public static double ComputeProtectionMinFastenerLength(double d, int t, IMaterialTimber timber, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, bool closedJoint, bool horizontal)
        {
            if (plasterboards.Count > 2 || plasterboards.Count <= 0) throw new Exception("The amount of plasterboard should be 1 or 2");
            if (!plasterboards.Contains(PlasterboardType.TypeF) && !plasterboards.Contains(PlasterboardType.TypeA)) throw new Exception("The amount of plasterboard should be 1 or 2");


            //Compute charring depth at tf

            double dchar;
            //Bn, Design notional charring rate under standard fire exposure 
            //tch, start of charring
            //tf, Failure time of protection 
            //ta, time limit above which the charring rate gets back to Bn
            double d0 = 7.0;
            double tch = ComputeCombustionStart(plasterboards, plasterboardthicknesses, closedJoint);
            double tf = ComputePanelFailureTime(tch, plasterboards, plasterboardthicknesses, horizontal);
            double K2 = ComputeK2(plasterboards, plasterboardthicknesses);
            double Bn = timber.Bn;
            double k0 = ComputeK0(tch, t, true);


            //Due to the implementation of different sources for the failure time of the protection (ÖNORM, Draft EN-1995-2 [2020])
            //The calculation of the protection failure time can be shorter than the combustion start.
            //In this case, we set tch = tf
            if (tf < tch) tch = tf;

            // t < tch
            if (t < tch) dchar = 0;

            // tch <= t <= ta
            else if (t >= tch && t <= tf) dchar =  K2 * Bn * (t - tch) + k0 * d0;

            else dchar = K2 * Bn * (tf - tch) + k0 * d0;


            double hp = plasterboardthicknesses.Sum();

            //The minimal length is considered as the panels thickness + the maximum of
            // - 6x fastener diameter
            // - The carbonised timber (when the protection fails) + 10mm according to EN 1995-1-2 §3.4.3.4 (4) - Eq 3.16
            return Math.Max(d * 6 + hp, dchar + 10 + hp);
        }

        /// <summary>
        /// Insulation coefficient according to DIN EN 1995-1-2 §3.4.3.2 (2)
        /// </summary>
        /// <param name="plasterboards">List of the type of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="plasterboardthicknesses">List of the thicknesses of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static double ComputeK2(List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses)
        {
            if (plasterboards.Count > 2 || plasterboards.Count <= 0) throw new Exception("The amount of plasterboard should be 1 or 2");
            if (plasterboards.Any(p => p != PlasterboardType.TypeF)) return 0;
            else return 1 - 0.018 * plasterboardthicknesses.Last();         //DIN EN 1995-1-2 Eq 3.7
        }

        /// <summary>
        /// Compute the time limit above which the charring rate gets back to Bn
        /// </summary>
        /// <param name="tch">CombustionStart</param>
        /// <param name="tf">PanelFailureTime</param>
        /// <param name="k2"></param>
        /// <param name="Bn"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static double ComputeTa(double tch, double tf, double k2, double Bn)
        {
            //Due to the implementation of different sources for the failure time of the protection (ÖNORM, Draft EN-1995-2 [2020])
            //The calculation of the protection failure time can be shorter than the combustion start.
            //In this case, we set tch = tf
            if (tf < tch) tch = tf;


            double ta = 0;
            if (tch == tf)
            {
                ta = Math.Min(2 * tf, 25 / (K3 * Bn) + tf);//DIN EN 1995-1-2 Eq 3.8
            }
            else
            {
                ta = (25 - (tf - tch) * k2 * Bn) / (K3 * Bn) + tf;  //DIN EN 1995-1-2 Eq 3.9
            }

            return ta;
        }

        public const int K3 = 2;//DIN EN 1995-1-2 §3.4.3.2 (4)

        /// <summary>
        /// Plasterboard type according to EN 520
        /// </summary>
        public enum PlasterboardType
        {
            TypeA,
            TypeF
        }


        #endregion

        #region helper functions
        /// <summary>
        /// Control function to ensure that IMaterialTimber interface is passsed as argument
        /// </summary>
        /// <param name="material">material object to test</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        internal static IMaterialTimber CheckMaterialTimber(IMaterial material)
        {
            if (!(material is IMaterialTimber)) throw new Exception("This method is currently only implemented for timber materials");
            return (IMaterialTimber)material;
        }

        /// <summary>
        /// Helper function to get a EC5_Utilities.ServiceClass Enum based on a string
        /// </summary>
        /// <param name="SC">Service Class</param>
        /// <returns></returns>
        public static ServiceClass GetServiceClass(string SC)
        {
            switch (SC)
            {
                case "SC1":
                    return ServiceClass.SC1;

                case "SC2":
                    return ServiceClass.SC2;

                case "SC3":
                    return ServiceClass.SC3;

                default:
                    throw new Exception("Service class should be defined as SC1, SC2 or SC3");
            }
        }

        /// <summary>
        /// Helper function to get a EC5_Utilities.ServiceClass Enum based on an integer
        /// </summary>
        /// <param name="SC">Service Class</param>
        /// <returns></returns>
        public static ServiceClass GetServiceClass(int SC)
        {
            switch (SC)
            {
                case 1:
                    return ServiceClass.SC1;

                case 2:
                    return ServiceClass.SC2;

                case 3:
                    return ServiceClass.SC3;

                default:
                    throw new Exception("Service class should be defined as 1, 2 or 3");
            }
        }

        /// <summary>
        /// Helper function to get a EC5_Utilities.LoadDuration Enum based on a string
        /// </summary>
        /// <param name="LD">LoadDuration</param>
        /// <returns></returns>
        public static LoadDuration GetLoadDuration(string LD)
        {

            switch (LD)
            {
                case "Permanent":
                    return LoadDuration.Permanent;

                case "LongTerm":
                    return LoadDuration.LongTerm;

                case "MediumTerm":
                    return LoadDuration.MediumTerm;

                case "ShortTerm":
                    return LoadDuration.ShortTerm;

                case "Instantaneous":
                    return LoadDuration.Instantaneous;

                case "ShortTerm_Instantaneous":
                    return LoadDuration.ShortTerm_Instantaneous;

                default:
                    throw new Exception("Load Duration should be defined as Permanent\nLongTerm\nMediumTerm\nShortTerm\nInstantaneous\nShortTerm_Instantaneous");
            }

        }

        #endregion


    }
}
