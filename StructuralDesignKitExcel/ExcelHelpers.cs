using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitExcel
{
    public class ExcelHelpers
    {

        /// <summary>
        /// Return a timber material object based on a string. Look into the database if t´his material exists. if not, an exception is thrown
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IMaterialTimber GetTimberMaterial(string material)
        {
            string ExceptionMaterialUnknown = string.Format("The material {0} is not part of the database, please define a correct material name or create a new material using the SDK.Material.Create function", material);

            //Lookup first letters to orient the search
            switch (material[0])
            {
                //Lookup if material is a defined Softwood
                case 'C':
                    if (Enum.GetNames(typeof(MaterialTimberSoftwood.Grades)).Contains(material))
                    {
                        return new MaterialTimberSoftwood(material);
                    }
                    else throw new Exception(ExceptionMaterialUnknown);

                //Lookup if material is a defined Hardwood
                case 'D':
                    if (Enum.GetNames(typeof(MaterialTimberHardwood.Grades)).Contains(material))
                    {
                        return new MaterialTimberHardwood(material);
                    }
                    else throw new Exception(ExceptionMaterialUnknown);


                //Lookup if material is a defined Glulam
                case 'G':
                    if (Enum.GetNames(typeof(MaterialTimberGlulam.Grades)).Contains(material))
                    {
                        return new MaterialTimberGlulam(material);
                    }
                    else if (Enum.GetNames(typeof(MaterialTimberBaubuche.Grades)).Contains(material))
                    {
                        return new MaterialTimberBaubuche(material);
                    }
                    else throw new Exception(ExceptionMaterialUnknown);

                default:
                    throw new Exception(ExceptionMaterialUnknown);
            }
        }

        public static CrossSectionRectangular CreateCrossSection (string crossSectionTag)
        {
            string error = "The cross section tag does not respect the defined syntax (i.e: CS_R_100x200_GL24h)";
            //CS_R_100x200_GL24h
            var CS = crossSectionTag.Split('_');
            IMaterialTimber material;

            int b = 0;
            int h = 0;
            
            if (CS.Length != 4) throw new Exception(error);

            if (CS[0] != "CS") throw new Exception(error);
            if (CS[1] != "R") throw new Exception("Currently only Rectangular cross-sections are supported");
            var bxh = CS[2].Split('x');
            if (bxh.Length != 2) throw new Exception(error);
            else
            {
                b = Int32.Parse(bxh[0]);
                h = Int32.Parse(bxh[1]);
            }

            try
            {
                material = GetTimberMaterial(CS[3]);
            }
            catch (Exception)
            {

                throw;
            }
            return new CrossSectionRectangular(b, h, material);
        }





        public static string CreateRectCrossSectionTag(double b, double h, IMaterialTimber material)
        {

            //CS_R_100x200_GL24h
            return String.Format("CS_R_{0}_{1}_{2}",
                b,
                h,
                material.Grade);
           

        }

    
    }
}
