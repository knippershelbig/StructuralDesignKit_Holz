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
using static StructuralDesignKitLibrary.EC5.EC5_Factors;
using StructuralDesignKitLibrary.CrossSections.Interfaces;

namespace StructuralDesignKit
{
    internal class Program
    {
        static void Main(string[] args)
        {



            var mat = new MaterialTimberHardwood(MaterialTimberHardwood.Grades.D35);
            var cs1 = new CrossSectionRectangular(50, 600, mat);


            //foreach (ServiceClass SC in Enum.GetValues(typeof(ServiceClass)))
            //{
            //    Console.WriteLine(SC.ToString());
            //    foreach (LoadDuration LC in Enum.GetValues(typeof(LoadDuration)))
            //    {
            //        Console.WriteLine(LC.ToString() + " -> " + EC5_Factors.Kmod(TimberType.Softwood, SC, LC).ToString());
            //        //Console.WriteLine(timber.ToString());
            //    }
            //}



            double kmod = Kmod(mat.Type, ServiceClass.SC1, LoadDuration.Permanent);
            double ym = Ym(mat.Type);



            Console.WriteLine(String.Format(
                "Tension -> {0:0.00}\n" +
                "Compression Y -> {1:0.00}\n" +
                "Bending  -> {2:0.00}\n" +
                "Shear  -> {3:0.00}\n" +
                "Torsion  -> {4:0.00}",
                EC5_CrossSectionCheck.TensionParallelToGrain(cs1.ComputeNormalStress(10), mat.Ft0k, kmod, ym, EC5_Factors.Kh_Tension(mat.Type, cs1.B), 1),
                EC5_CrossSectionCheck.CompressionParallelToGrain(cs1.ComputeNormalStress(10), mat.Fc0k, kmod, ym),
                EC5_CrossSectionCheck.Bending(cs1.ComputeStressBendingY(10), cs1.ComputeStressBendingZ(0.5), mat.Fmyk, mat.Fmzk, cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), Kh_Bending(mat.Type, cs1.B)),
                EC5_CrossSectionCheck.Shear(cs1.ComputeShearY(5), cs1.ComputeShearZ(10), mat.Fvk, mat, kmod, ym),
                EC5_CrossSectionCheck.Torsion(cs1.ComputeTorsion(0.1),cs1.ComputeShearY(5), cs1.ComputeShearZ(10), mat.Fvk,cs1, mat, kmod, ym)));


            //cs1.ComputeShearY(10).ToString(),
            //cs1.ComputeShearZ(10).ToString(),
            //cs1.ComputeTorsion(10).ToString()));

            Console.ReadLine();
        }
    }

}

