using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.IdentityModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.EC5
{
    static public class EC5_ConnectionCheck
    {

        //Block shear and plug shear failure at multiple dowel-type steel-to-timber connections
        static public double BlockShearFailureCapacity(ISteelTimberShear connection, double SumLti, double SumLvi)
        {


            double tef = 0;

            string failureMode = connection.FailureMode;
            double Anet_v;

            IMaterialTimber timber = connection.Timber;

            switch (failureMode)
            {

                case "c":
                case "f":
                case "j":
                case "l":
                case "k":
                case "m":
                    tef = connection.TimberThickness;
                    break;

                case "a":
                    tef = 0.4 * connection.TimberThickness;
                    break;

                case "b":
                    tef = 1.4 * Math.Sqrt(connection.Fastener.MyRk / (connection.Fastener.Fhk * connection.Fastener.Diameter));
                    break;

                case "e":
                case "h":
                    tef = 2 * Math.Sqrt(connection.Fastener.MyRk / (connection.Fastener.Fhk * connection.Fastener.Diameter));
                    break;

                case "d":
                case "g":
                    tef = connection.TimberThickness * (Math.Sqrt(2 + 4 * connection.Fastener.MyRk / (connection.Fastener.Fhk * connection.Fastener.Diameter * Math.Pow(connection.TimberThickness, 2))) - 1);
                    break;
            }

            Anet_v = ComputeAnet_v(failureMode, tef, SumLvi, SumLti);

            double Anet_t = ComputeAnet_t(connection.FailureMode,connection.TimberThickness, SumLti );

            double tensileStrength = 1.5 * Anet_t * timber.Ft0k;
            double shearStrength = 0.7 * Anet_v * timber.Fvk;


            return Math.Max(tensileStrength, shearStrength);


        }

        static private double ComputeAnet_t(string failureMode, double t, double sumLnetT)
        {
            switch (failureMode)
            {

                case "a":
                case "b":
                case "c":
                case "d":
                case "e":
                case "j":
                case "l":
                case "k":
                case "m":
                    return t * sumLnetT;

                case "f":
                case "g":
                case "h":
             
                    return t * sumLnetT*2;

                default: throw new Exception("Failure mode not considered for BlockShearFailureCapacity");
            }
            


        }

        static private double ComputeAnet_v(string failureMode, double tef, double sumLnetV, double sumLnetT)
        {
            switch (failureMode)
            {

                case "c":
                    return tef * sumLnetV;

                case "f":
                case "j":
                case "l":
                case "k":
                case "m":
                    return tef * sumLnetV * 2;

                case "a":
                case "b":
                case "d":
                case "e":
                    return sumLnetV / 2 * (sumLnetT + 2 * tef);

                case "h":
                case "g":
                    return sumLnetV / 2 * (sumLnetT + 2 * tef) * 2;


                default: throw new Exception("Failure mode not considered for BlockShearFailureCapacity");
            }



        }


    }
}
