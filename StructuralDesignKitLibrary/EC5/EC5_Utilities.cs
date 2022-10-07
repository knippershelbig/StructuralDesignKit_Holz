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

        #endregion


    }
}
