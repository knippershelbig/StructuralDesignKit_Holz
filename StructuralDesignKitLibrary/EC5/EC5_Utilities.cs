using StructuralDesignKitLibrary.Materials;
using System;
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
