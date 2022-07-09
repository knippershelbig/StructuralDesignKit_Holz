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


    }
}
