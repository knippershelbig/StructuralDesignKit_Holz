using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.EC5;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;

namespace StructuralDesignKit
{
    internal class Program
    {
        static void Main(string[] args)
        {



            var mat = new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24);
            var cs1 = new CrossSectionRectangular(100, 200, mat);


            foreach (ServiceClass SC in Enum.GetValues(typeof(ServiceClass)))
            {
                Console.WriteLine(SC.ToString());
                foreach (LoadDuration LC in Enum.GetValues(typeof(LoadDuration)))
                {
                    Console.WriteLine(LC.ToString() + " -> " + EC5_Factors.Kmod(TimberType.Softwood, SC, LC).ToString());
                    //Console.WriteLine(timber.ToString());
                }
            }

            Console.ReadLine();
        }



    }

}
