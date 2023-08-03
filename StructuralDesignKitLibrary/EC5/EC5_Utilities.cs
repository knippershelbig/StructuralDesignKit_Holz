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
        //Charring calculation for beams
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
        /// Return the charring depth in mm of a protected beam. Result is per exposed face
        /// </summary>
        /// <param name="t">exposed duration</param>
        /// <param name="timber">Timber member material</param>
        /// <param name="plasterboards">List of the type of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="plasterboardthicknesses">List of the thicknesses of plasterboards (max 2); For 2 plasterboards, the first one is the external one</param>
        /// <param name="closedJoint">if true, joints between plasterboards is considered <2mm; if false > 2mm</param>
        /// <param name="horizontal">Boolean value to represent if the protection is horizontal or vertical</param>
        /// <returns></returns>
        public static double ComputeCharringDepthProtectedBeam(int t, IMaterialTimber timber, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, bool closedJoint, bool horizontal, int fastenerLength)
        {
            if (plasterboardthicknesses.Count > 2 || plasterboards.Count > 2) throw new Exception("Standardized protection in the EN 1995-1-2 §3.4.3.3 accounts only for a maximum of 2 boards");

            //Bn, Design notional charring rate under standard fire exposure 
            //tch, start of charring
            //tf, Failure time of protection 
            //ta, time limit above which the charring rate gets back to Bn
            double d0 = 7.0;
            double tch = ComputeCombustionStart(plasterboards, plasterboardthicknesses, closedJoint);

            double tf = ComputePanelFailureTime(tch, plasterboards, plasterboardthicknesses, fastenerLength, horizontal, timber);

            //
            //if (tf < tch) tch = tf;

            double K2 = ComputeK2(plasterboards, plasterboardthicknesses);

            double Bn = timber.Bn;

            double ta = ComputeTa(tch, tf, K2, Bn);

            double k0 = ComputeK0(tch, t, true);

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
        /// Compute the Panel Failure Time according to the Austrian approach
        /// </summary>
        /// <param name="tch"></param>
        /// <param name="plasterboards"></param>
        /// <param name="plasterboardthicknesses"></param>
        /// <param name="fastenerLenght"></param>
        /// <param name="horizontal">Boolean value to represent if the protection is horizontal or vertical</param>
        /// <param name="timber"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static double ComputePanelFailureTime(double tch, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, int fastenerLenght, bool horizontal, IMaterialTimber timber)
        {
            if (plasterboards.Count > 2 || plasterboards.Count <= 0) throw new Exception("The amount of plasterboard should be 1 or 2");
            if (!plasterboards.Contains(PlasterboardType.TypeF) && !plasterboards.Contains(PlasterboardType.TypeA)) throw new Exception("The amount of plasterboard should be 1 or 2");


            //If the plasterboards are not type F, tf = tch
            if (plasterboards.Contains(PlasterboardType.TypeA)) return tch;


            //---------------------------------------------
            //Max value for panel fasteners failure according to //DIN EN 1995-1-2 §C2.3
            //---------------------------------------------


            int lf = fastenerLenght;
            int la_min = 10;
            double hp = plasterboardthicknesses.Sum();
            double ks = 1.1; // DIN EN 1995-1-2 table C1 - Value taken for a minimal cross section of 60mm; Higher value might apply in case of timberframe wall (38 & 45mm); Unlikely for structural element
            double k2 = 0;
            double kj = 0;

            //We can consider that there will always be a joint, so onJoint = true
            bool OnJoint = true;
            if (OnJoint)
            {
                k2 = 0.86 - 0.0037 * hp;    //DIN EN 1995-1-2 Eq C4
                kj = 1.15;                  //DIM EN 1995-1-2 Eq C11
            }
            else //No plasterboard joint over the considered area - implemented but not used
            {
                k2 = 1.05 - 0.0073 * hp;    //DIN EN 1995-1-2 Eq C3
                kj = 1;                     //DIM EN 1995-1-2 Eq C10
            }


            double kn = 1.5;                //DIN EN 1995-1-2 §C2.1 (1)

            double tf_f = tch + (lf - la_min - hp) / (ks * k2 * kn * kj * timber.B0); //DIN EN 1995-1-2 Eq C9

            if (tf_f < 0) throw new Exception("Fasteners are too short - they should at least go 10mm in the uncharred timber");




            //--------------------------------------------------------------------------------------------------------------------------------
            //As the current version of the Eurocode 5 does not provide a mean to calculate these value, the Austrian approach based on the 
            //results from the Holzforschung Austria [Teibinger und Matzinger 2010] is implemented
            //--------------------------------------------------------------------------------------------------------------------------------

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

            //making sure failure time of the panel is not lower than start of charring time 
            //This can happen with the austrian approach and this lead to panel type A performing better than panel type F.
            //To our opinion, this should not be the case, hence the decision to put a lower threshold to the panel failure time
            if (tf_p < tch) tf_p = tch;

            return (Math.Min(tf_f, tf_p));
        }


        public static double ComputePanelFailureTimeDepreciated(double tch, List<PlasterboardType> plasterboards, List<double> plasterboardthicknesses, int fastenerLenght, bool OnJoint, IMaterialTimber timber)
        {

            if (plasterboards.Count > 2 || plasterboards.Count <= 0) throw new Exception("The amount of plasterboard should be 1 or 2");
            foreach (PlasterboardType plasterboard in plasterboards)
            {
                if (plasterboard != PlasterboardType.TypeF && plasterboard != PlasterboardType.TypeA) throw new Exception("Plasterboard type not implemented");
                if (plasterboard == PlasterboardType.TypeA) return tch;
            }



            //Max value for panel fasteners failure according to //DIN EN 1995-1-2 §C2.3
            int lf = fastenerLenght;
            int la_min = 10;
            double hp = plasterboardthicknesses.Sum();
            double ks = 1.1; // DIN EN 1995-1-2 table C1 - Value taken for a minimal cross section of 60mm; Higher value might apply in case of timberframe wall (38 & 45mm); Unlikely for structural element
            double k2 = 0;
            double kj = 0;

            //We can consider that there will always be a joint, so onJoint = true
            OnJoint = true;
            if (OnJoint)
            {
                k2 = 0.86 - 0.0037 * hp;    //DIN EN 1995-1-2 Eq C4
                kj = 1.15;                  //DIM EN 1995-1-2 Eq C11
            }
            else //No plasterboard joint over the considered area
            {
                k2 = 1.05 - 0.0073 * hp;    //DIN EN 1995-1-2 Eq C3
                kj = 1;                     //DIM EN 1995-1-2 Eq C10
            }



            double kn = 1.5;                //DIN EN 1995-1-2 §C2.1 (1)


            double tf_f = tch + (lf - la_min - hp) / (ks * k2 * kn * kj * timber.B0); //DIN EN 1995-1-2 Eq C9

            if (tf_f < 0) throw new Exception("Fasteners are too short - they should at least go 10mm in the uncharred timber");

            //-------------------------------------------------------------------------------------------------------------
            //As the current version of the Eurocode 5 does not provide a mean to calculate these value, it was currently
            //chosen to implement the equations provided by the Draft of the new version of the Eurocode 5, even if it is 
            //not yet implemented.
            //Conservatively, only the value for horizontal time failure have been implemented and only for type F panels
            //The whole process according to this standard will be implemented when the new Eurocode will be enforced ~2025
            //-------------------------------------------------------------------------------------------------------------


            double tf_p = 0;

            //Max value for panel failure according to //EN 1995-1-2:2020 (E) Draft 2021  - §5.4.2.3 - Table 5.4
            if (plasterboards.Count == 1)
            {
                //verify conformity with standard

                if (plasterboardthicknesses[0] > 18 || plasterboardthicknesses[0] < 9) throw new Exception("Single plasterboard thickness needs to be comprised between 9 and 18mm");


                if (plasterboards[0] == PlasterboardType.TypeF) tf_p = 1.3 * hp + 9;
                else if (plasterboards[0] == PlasterboardType.TypeA) tf_p = 2.1 * hp - 9;
                else throw new Exception("Plasterboard type not implemented");

                if (tf_p < tch) tf_p = tch; //making sure failure time of the panel is not lower than start of charring time | To ensure compatibility between Draft of the new EC5 and current EC5

            }
            else
            {
                //verify conformity with standard
                double thick = plasterboardthicknesses.Sum();
                if (thick > 36 || thick < 24) throw new Exception("the overall plasterboards thickness needs to be comprised between 24 and 36mm");


                if (plasterboards[0] == PlasterboardType.TypeF) tf_p = 1.5 * hp + 15;
                else if (plasterboards[0] == PlasterboardType.TypeA) tf_p = 1.7 * hp - 13;

                else throw new Exception("Plasterboard type not implemented");

                tf_p *= 1.1;
                if (tf_p < tch) tf_p = tch;//making sure failure time of the panel is not lower than start of charring time | To ensure compatibility between Draft of the new EC5 and current EC5
            }



            return (Math.Min(tf_f, tf_p));

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
            double ta = 0;
            if (tch == tf)
            {
                ta = Math.Min(2 * tf, 25 / (K3 * Bn) + tf);//DIN EN 1995-1-2 Eq 3.8
            }
            else if (tch < tf)
            {
                ta = (25 - (tf - tch) * k2 * Bn) / (K3 * Bn) + tf;  //DIN EN 1995-1-2 Eq 3.9
            }
            else throw new Exception("Case not covered");

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
