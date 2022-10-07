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



            var mat = new MaterialTimberSoftwood(MaterialTimberSoftwood.Grades.C24);
            var cs1 = new CrossSectionRectangular(100, 1200, mat);




            double kmod = Kmod(mat.Type, ServiceClass.SC1, LoadDuration.Permanent);
            double ym = Ym(mat.Type);

            double N = -150;
            double Vy = 0;
            double Vz = 0;
            double Mx = 0;
            double My = 100;
            double Mz = 0;

            double leffy = 0;
            double leffz = 0;
            double ltb_Eff = 5000;


            var loads = new List<double>() { My, Mz, N };

            Console.WriteLine(String.Format(
                "Tension -> {0:0.00}\n" +
                "Compression Y -> {1:0.00}\n" +
                "Bending  -> {2:0.00}\n" +
                "Shear  -> {3:0.00}\n" +
                "Torsion  -> {4:0.00}\n" +
                "Bending and tension  -> {5:0.00}\n" +
                "Bending and compression -> {6:0.00}\n" +
                "Bending and buckling -> {7:0.00}\n" +
                "Lateral Torsional Buckling -> {8:0.00}",
                EC5_CrossSectionCheck.TensionParallelToGrain(cs1.ComputeNormalStress(N), mat, kmod, ym, EC5_Factors.Kh_Tension(mat.Type, cs1.B), 0.1),
                EC5_CrossSectionCheck.CompressionParallelToGrain(cs1.ComputeNormalStress(N), mat, kmod, ym),
                EC5_CrossSectionCheck.Bending(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
                EC5_CrossSectionCheck.Shear(cs1.ComputeShearY(Vy), cs1.ComputeShearZ(Vz), mat, kmod, ym),
                EC5_CrossSectionCheck.Torsion(cs1.ComputeTorsion(Mx), cs1.ComputeShearY(Vy), cs1.ComputeShearZ(Vz), cs1, mat, kmod, ym),
                EC5_CrossSectionCheck.BendingAndTension(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1, Kh_Tension(mat.Type, cs1.B)),
                EC5_CrossSectionCheck.BendingAndCompression(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
                EC5_CrossSectionCheck.BendingAndBuckling(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), leffy, leffz, cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
                EC5_CrossSectionCheck.LateralTorsionalBuckling(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), leffy, leffz, ltb_Eff, cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1)));


            double angle = 10;
            Console.WriteLine("Compression at an angle {0}: {1:0.00}", angle,
                EC5_CrossSectionCheck.CompressionAtAnAngleToGrain(2, angle, cs1.Material, kmod, ym, 1.5));


            Console.WriteLine(EC5_Factors.Kp(320, 25, 25000));





            DescriptionAttribute attribute = (DescriptionAttribute)typeof(EC5_Factors).GetMethod("Kdis").GetCustomAttributes(false)[0];
            
            //var description = (DescriptionAttribute)attribute;
            //var text = description.Description;


 

            Console.ReadLine();
        }
    }

}

