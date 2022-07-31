using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;
using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;


namespace StructuralDesignKitExcel
{


    //public static class MyFunctions
    //{

    //    [ExcelFunction(Description = "My first .NET function")]
    //    public static string SayHello(string name)
    //    {
    //        return "Hello " + name;
    //    }

    //    [ExcelFunction(Description = "Return a SDK Material")]
    //    public static Object material(string materialName)
    //    {
    //        var mat = new MaterialTimberGlulam(materialName);
    //        var guid = Guid.NewGuid();
    //        return mat.Ft0k;
    //    }

    //    [ExcelFunction(Description = "Return a FCK from a material")]
    //    public static double fc0k(object material)
    //    {
    //        var mat = (IMaterialTimber)material;
    //        return mat.Fc0k;
    //    }

    //    //        // Get the type corresponding to the class MyClass.
    //    //Type myType = typeof(MyClass1);
    //    //// Get the object of the Guid.
    //    //Guid myGuid = (Guid)myType.GUID;
    //    //Console.WriteLine("The name of the class is "+myType.ToString());
    //    //Console.WriteLine("The ClassId of MyClass is "+myType.GUID);	


    //    [ExcelFunction("Returns the answer")]
    //    public static object MyFunction([ExcelArgument("The unimportant input")] object input)
    //    {
    //        return 42;
    //    }







    //}



    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            try
            {
                ExcelDna.IntelliSense.IntelliSenseServer.Install();
            }

            catch (Exception e)
            {

            }
        }


        public void AutoClose()
        {
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();
        }




    }

    public static class myfunction
    {
        [ExcelFunction("Returns the answer")]
        public static object MyFunction([ExcelArgument("The unimportant input")] object input)
        {
            return 42;
        }


        [ExcelFunction(
           Description =
               "Returns the COmpression perpenducular to the grain of a material.",
           Name = "SDK.Materials.Fmk",
           Category = "SDK.Materials",
           IsHidden = false)]
        public static double TestFunction(
           [ExcelArgument(Description = "This is my first test function")]
            string name)
        {
            return MaterialTimberSoftwood.fc0k[name];
        }


        //[ExcelFunction(Description = "Find a material based on a string",
        //    Name = "SDK.Material.GetMaterial",
        //    IsHidden = false,
        //    Category = "SDK.EC5_Factors")]
        //public static double buckling()
        //{
        //    var Kcs = EC5_Factors.Kc()
        //}


        [ExcelFunction(Description = "Find a material based on a string",
       Name = "SDK.Material.GetMaterial",
       IsHidden = false,
       Category = "SDK.EC5_Materials")]
        public static string Material(string material)
        {
            IMaterialTimber mat = ExcelHelpers.GetTimberMaterial(material);
            return mat.Grade;

        }

        [ExcelFunction(Description = "TestFunction",
Name = "SDK.Material.Test",
IsHidden = false,
Category = "SDK.EC5_Materials")]
        public static string CrossSection(string CrossSectionTag)
        {
            CrossSectionRectangular CS = ExcelHelpers.CreateCrossSection(CrossSectionTag);
            return String.Format("{0}x{1}mm",CS.B.ToString(), CS.H.ToString());

        }



        //material encryypt and decript


        //Rectangular Cross section encrypt Decript
        //CS_R_100x200_GL24h
        //CS_C_120_C24


        //Expose all checks
        //Check all at once given a load torsor


        //Create drop down list with materials

        //Create a list of all material properties in excel (kind of deconstruct material)

        //Create a material which can be used following the IMaterial Timber Interface

        //Expose drop down Material type (baubuche, ...)

        //define material type based on names Find material type 

    }

}






