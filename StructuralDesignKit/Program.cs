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
using System.IO;
using System.Xml;
using StructuralDesignKitLibrary.Connections.Fasteners;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using StructuralDesignKitLibrary.Connections.TimberTimberShear;
using StructuralDesignKitLibrary.Utilities;
using StructuralDesignKitLibrary.RFEM;

namespace StructuralDesignKit
{
    internal class Program
    {
        static void Main(string[] args)
        {

            
     

            //Console.WriteLine("Fastener checks\n");

            //var Dowel = new FastenerDowel(10, 360);
            //var bolt = new FastenerBolt(20, 800);
            ////var timber = new MaterialTimberHardwood(MaterialTimberHardwood.Grades.D30);
            //var timber = new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL28h);
            //double angle1 = 43;

            //var shearCapacity2 = new SteelSingleInnerPlate(Dowel,6, angle1, timber, 45, false);
            ////var shearCapacity2 = new TimberTimberSingleShear(bolt,timber, 160, 90,timber,175,0,true);

            //Dowel.ComputeSpacings(angle1);

            //for (int i = 0; i < shearCapacity2.Capacities.Count; i++)
            //{
            //Console.WriteLine(String.Format("Failure mode {0} = {1:0}N", shearCapacity2.FailureModes[i], shearCapacity2.Capacities[i]));

            //}

            //Console.WriteLine(String.Format("Shear Capacity = {0:0.00}KN (Failure mode {1})", shearCapacity2.Capacity/1000,shearCapacity2.FailureMode));
            //Console.WriteLine(String.Format("a1={0:0}mm\na2={1:0}mm\na3t={2:0}mm\na3c={3:0}mm\na4t={4:0}mm\na4c={5:0.}mm\n",
            //    Dowel.a1min, Dowel.a2min, Dowel.a3tmin, Dowel.a3cmin, Dowel.a4tmin, Dowel.a4cmin));

            //Console.WriteLine(SDKUtilities.LinearInterpolation(5, 10, 2, 25, 52));



            //List<String> propertiesToModify = new List<string> { "Grade", "Fmyk"};
            //List<Object> values = new List<Object> { "GL24h_Modified", "32" };


            ////MaterialTimberGeneric genericTimber = new MaterialTimberGeneric(new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h),propertiesToModify,values);
            //var mat = new MaterialTimberGeneric(new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h), propertiesToModify, values);

            //var cs1 = new CrossSectionRectangular(100, 1200, mat);

            //double kmod = Kmod(mat.Type, ServiceClass.SC1, LoadDuration.Permanent);
            //double ym = Ym(mat.Type);

            //double N = -150;
            //double Vy = 0;
            //double Vz = 0;
            //double Mx = 0;
            //double My = 100;
            //double Mz = 0;

            //double leffy = 0;
            //double leffz = 0;
            //double ltb_Eff = 5000;


            //var loads = new List<double>() { My, Mz, N };

            //Console.WriteLine(String.Format(
            //    "Tension -> {0:0.00}\n" +
            //    "Compression Y -> {1:0.00}\n" +
            //    "Bending  -> {2:0.00}\n" +
            //    "Shear  -> {3:0.00}\n" +
            //    "Torsion  -> {4:0.00}\n" +
            //    "Bending and tension  -> {5:0.00}\n" +
            //    "Bending and compression -> {6:0.00}\n" +
            //    "Bending and buckling -> {7:0.00}\n" +
            //    "Lateral Torsional Buckling -> {8:0.00}",
            //    EC5_CrossSectionCheck.TensionParallelToGrain(cs1.ComputeNormalStress(N), mat, kmod, ym, EC5_Factors.Kh_Tension(mat.Type, cs1.B), 0.1),
            //    EC5_CrossSectionCheck.CompressionParallelToGrain(cs1.ComputeNormalStress(N), mat, kmod, ym),
            //    EC5_CrossSectionCheck.Bending(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
            //    EC5_CrossSectionCheck.Shear(cs1.ComputeShearY(Vy), cs1.ComputeShearZ(Vz), mat, kmod, ym),
            //    EC5_CrossSectionCheck.Torsion(cs1.ComputeTorsion(Mx), cs1.ComputeShearY(Vy), cs1.ComputeShearZ(Vz), cs1, mat, kmod, ym),
            //    EC5_CrossSectionCheck.BendingAndTension(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1, Kh_Tension(mat.Type, cs1.B)),
            //    EC5_CrossSectionCheck.BendingAndCompression(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
            //    EC5_CrossSectionCheck.BendingAndBuckling(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), leffy, leffz, cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1),
            //    EC5_CrossSectionCheck.LateralTorsionalBuckling(cs1.ComputeStressBendingY(My), cs1.ComputeStressBendingZ(Mz), cs1.ComputeNormalStress(N), leffy, leffz, ltb_Eff, cs1, mat, kmod, ym, Kh_Bending(mat.Type, cs1.H), 1)));


            Console.WriteLine("\n-------------------------------------------------------------------\n");

            var watch = System.Diagnostics.Stopwatch.StartNew();
            var model = RFEM_Utilities.OpenModel();
            watch.Start();
            //var res1 = RFEM_Utilities.GetAllStandardizedDisplacement(model,1);
            watch.Stop();
            Console.WriteLine("Algo1 = " + watch.ElapsedMilliseconds.ToString());
            watch.Reset();
            watch.Start();
            var res2 = RFEM_Utilities.GetAllStandardizedDisplacement(model,1);
            Console.WriteLine("Algo1 = " + watch.ElapsedMilliseconds.ToString());
            watch.Reset();
            RFEM_Utilities.CloseRFEMModel(model);




            Console.ReadLine();
        }
    }

}

