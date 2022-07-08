using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using StructuralDesignKitLibrary.Materials;
using StructuralDesignKitLibrary.CrossSections;

namespace StructuralDesignKit
{
    internal class Program
    {
        static void Main(string[] args)
        {



            var mat = new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24);
            var cs1 = new CrossSectionRectangular(100, 200, mat);

            Console.WriteLine(String.Format(
                "Area = {0:0}mm²\n" +
                "Inertia Y = {1:0}mm4\n" +
                "Inertia Z = {2:0}mm4\n" +
                "Section Modulus Y = {3:0}mm3\n" +
                "Section Modulus Z = {4:0}mm3",
                cs1.Area,
                cs1.MomentOfInertia_Y,
                cs1.MomentOfInertia_Z,
                cs1.SectionModulus_Y,
                cs1.SectionModulus_Z));

            Console.ReadLine();
        }



    }

}
