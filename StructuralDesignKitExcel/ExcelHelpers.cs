using Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.CrossSections;
using StructuralDesignKitLibrary.EC5;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace StructuralDesignKitExcel
{
    public static class ExcelHelpers
    {

        /// <summary>
        /// Return a timber material object based on a string. Look into the database if t´his material exists. if not, an exception is thrown
        /// </summary>
        /// <param name="material"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IMaterialTimber GetTimberMaterialFromTag(string material)
        {
            string ExceptionMaterialUnknown = string.Format("The material {0} is not part of the database, please define a correct material name or create a new material using the SDK.Material.Create function", material);


            if (material.Split('|').ToList().Count > 1 && material.Split('|')[0].Substring(0, 4) == "Base")
            {
                return CreateGenericMaterialFromTag(material);
            }


            else
            {
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

        }

        public static IMaterialTimber CreateGenericMaterialFromTag(string material)
        {
            var properties = material.Split('|').ToList();

            //If based on a existing material
            List<string> propertiesToModify = new List<string>();
            List<object> values = new List<object>();

            if (properties[0].Substring(0, 4) == "Base")
            {
                for (int i = 1; i <= properties.Count - 1; i++)
                {
                    propertiesToModify.Add(properties[i].Split('-')[0]);
                    values.Add(properties[i].Split('-')[1]);
                }
            }
            IMaterialTimber baseMaterial = GetTimberMaterialFromTag((string)properties[0].Split('-')[1]);
            return new MaterialTimberGeneric(baseMaterial, propertiesToModify, values, material);

        }

        public static CrossSectionRectangular CreateRectangularCrossSection(string crossSectionTag)

        {
            string error = "The cross section tag does not respect the defined syntax (i.e: CS_R_100x200_GL24h)";
            var CS = crossSectionTag.Split('_');
            IMaterialTimber material;

            int b = 0;
            int h = 0;

            if (CS.Length != 4)
            {
                if (CS[3] != "GL75h") throw new Exception(error);
                else CS[3] += "_" + CS[4];
            }

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
                material = GetTimberMaterialFromTag(CS[3]);
            }
            catch (Exception)
            {

                throw;
            }
            return new CrossSectionRectangular(b, h, material);
        }


        public static string CreateRectangularCrossSectionTag(double b, double h, IMaterialTimber material)
        {
            //CS_R_100x200_GL24h
            return String.Format("CS_R_{0}x{1}_{2}",
                b,
                h,
                material.Grade);
        }


        public static List<string> AllMaterialAsList()
        {
            List<string> allTimber = new List<string>();
            allTimber.AddRange(ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberSoftwood.Grades>());
            allTimber.AddRange(ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberHardwood.Grades>());
            allTimber.AddRange(ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberGlulam.Grades>());
            allTimber.AddRange(ExcelHelpers.GetStringValuesFromEnum<StructuralDesignKitLibrary.Materials.MaterialTimberBaubuche.Grades>());

            return allTimber;
        }


        public static List<string> GetStringValuesFromEnum<T>()
        {
            List<string> valuesString = new List<string>();

            var values = Enum.GetValues(typeof(T)).Cast<T>();
            foreach (var val in values)
            {
                valuesString.Add(val.ToString());
            }
            return valuesString;
        }


        //Check if a workbook is open - If not, open a blank WB
        public static void WorkBookOpen(Excel.Application app)
        {
            if (app.Workbooks.Count == 0)
            {
                app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                // app.Workbooks[0].Activate();
            }
        }


        /// <summary>
        /// Method to check if a fastener type is available in the SDK based on a string
        /// </summary>
        /// <param name="fastenerType"></param>
        /// <returns></returns>
        public static bool IsFastener(string fastenerType)
        {
            //Get fastener types available in SDK
            var availableFastenerTypes = Enum.GetNames(typeof(EC5_Utilities.FastenerType)).ToList();
            List<string> availableFastenerTypesLower = new List<string>();
            foreach (var item in availableFastenerTypes)
            {
                availableFastenerTypesLower.Add(item.ToLower());
            }
            if (availableFastenerTypesLower.Contains(fastenerType.ToLower())) return true;
            else return false;
        }


        /// <summary>
        /// Return a IFastener object based on a string, diameter and steel tensile strength
        /// </summary>
        /// <param name="fastenerType"></param>
        /// <param name="diameter"></param>
        /// <param name="fuk"></param>
        /// <returns></returns>
        public static IFastener GetFastener(string fastenerType, double diameter, double fuk)
        {
            //Get fasterner types in SDK
            var fasteners = Assembly.Load("StructuralDesignKitLibrary").GetTypes().Where(p => p.FullName.StartsWith("StructuralDesignKitLibrary.Connections.Fasteners")).ToList();

            //Get the fastener SDK type requested
            Type type = fasteners.Where(p => p.Name.ToLower().Contains(fastenerType.ToLower())).First();

            //Create an instance from the particular fastener type to compute the properties
            IFastener fastener = (IFastener)Activator.CreateInstance(type, diameter, fuk);

            return fastener;
        }


        //Create BoltTag
        public static string GenerateBoltTag(double diameter, double Fuk)
        {
            return string.Format("Bolt_D{0}_Fu{1}", diameter, Fuk);
        }

        //Create DowelTag
        public static string GenerateDowelTag(double diameter, double Fuk)
        {
            return string.Format("Dowel_D{0}_Fu{1}", diameter, Fuk);
        }


        public static IFastener GetFastenerFromTag(string fastenerTag)
        {
            var fastener = fastenerTag.Split('_');
            string error = "Format not recognized";
            //Proofing
            if (string.IsNullOrWhiteSpace(fastenerTag)) throw new Exception("FastenerTag is nullOrWhiteSpace");
            if (!IsFastener(fastener[0])) throw new Exception("Fastener type not recognized. (Bolt_D10_Fu800 | Dowel_D10_Fu360");
            if (fastener[1][0] != 'D') throw new Exception(error);


            string type = fastener[0];
            double diameter = Convert.ToDouble(fastener[1].Remove(0, 1));
            double fu = Convert.ToDouble(fastener[2].Remove(0, 2));

            return GetFastener(type, diameter, fu);
        }


        //Create DowelTag

        //Bolt_D10_Fu800
        //
        //Screw_Dint6_Dext10_Fu1000_Lg_30_Faxk11.2_Rhoa350
        //Nail_DDext10_Fu1000_Lg_30_Faxk11.2_Rhoa350

    }
}
