using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Dlubal.RFEM5;
using Dlubal.DynamPro;
using System.Security.Claims;
using System.Reflection;

namespace StructuralDesignKitLibrary.RFEM
{
    public static class RFEM_Utilities
    {
        public static IModel GetActiveModel()
        {
            IModel model = null;
            try
            {
                // gets interface to an opened RFEM model
                model = Marshal.GetActiveObject("RFEM5.Model") as IModel;
                // checks RF-COM license and locks the application for using by COM
                model.GetApplication().LockLicense();




                // unlocks the application and releases RF-COM license

                model.GetApplication().UnlockLicense();

                // releases COM object
                model = null;
                // cleans Garbage Collector for releasing all COM interfaces and objects
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();

            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }

            return model;
        }

        public static IModel OpenModel()
        {
            IModel model = null;
            try
            {
                // gets interface to an opened RFEM model
                model = Marshal.GetActiveObject("RFEM5.Model") as IModel;
                // checks RF-COM license and locks the application for using by COM
                model.GetApplication().LockLicense();
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }

            return model;
        }

        public static void CloseRFEMModel(IModel model)
        {
            try
            {
                // unlocks the application and releases RF-COM license

                model.GetApplication().UnlockLicense();

                // releases COM object
                model = null;
                // cleans Garbage Collector for releasing all COM interfaces and objects
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        /// <summary>
        /// Get the modal masses from RFEM as list (one per mode)
        /// </summary>
        /// <param name="model">RFEM Model</param>
        /// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
        /// <returns>Returns a list of double representing the modal masses</returns>
        public static List<double> GetModalMasses(IModel model, int NVC)
        {
            var dynPro = model.GetModule("DynamPro") as IDynamModule;

            IDynamResults dynamicResults = dynPro.GetResults();

            var modalMasses = dynamicResults.GetAllNVCMassFactors(NVC);

            List<double> modalMassesList = new List<double>();

            foreach (var mode in modalMasses)
            {
                modalMassesList.Add(mode.ModalMassMi);
            }

            return modalMassesList;
        }

        /// <summary>
        /// Get the natural frequencies calculated in RF-DynamPro
        /// </summary>
        /// <param name="model">RFEM Model</param>
        /// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
        /// <returns>Returns a list of double representing the natural frequencies calculated in RF-DynamPro</returns>
        public static List<double> GetNaturalFrequencies(IModel model, int NVC)
        {
            var DynamPro = model.GetModule("DynamPro") as IDynamModule;

            IDynamResults dynamicResults = DynamPro.GetResults();

            var naturalFrequencies = dynamicResults.GetAllNVCNaturalFrequencies(NVC);

            List<double> naturalFrequenciesList = new List<double>();

            foreach (var mode in naturalFrequencies)
            {
                naturalFrequenciesList.Add(mode.NaturalFrequency);
            }

            return naturalFrequenciesList;
        }




        /// <summary>
        /// Get all the standardized displacement for the different mode shapes and FE nodes
        /// </summary>
        /// <param name="model">RFEM Model</param>
        /// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
        /// <returns></returns>
        public static List<ModeShape> GetAllStandardizedDisplacementSLOW(IModel model, int NVC)
        {

            var DynamPro = model.GetModule("DynamPro") as IDynamModule;

            IDynamResults dynamicResults = DynamPro.GetResults();

            int nbModeShape = DynamPro.GetData().GetNVCParams(NVC).NumberModeShapes;
            var MeshNodes = model.GetCalculation().GetFeMesh().GetNodes();
            int nbMeshNode = MeshNodes.Length;

            List<ModeShape> modeShapes = new List<ModeShape>();

            for (int i = 0; i < nbMeshNode; i++)
            {
                ModeShape modeshape = new ModeShape(MeshNodes[i].No);
                for (int j = 0; j < nbModeShape; j++)
                {
                    var result = dynamicResults.GetNVCModeShapesOnFEPoint(NVC, j + 1, MeshNodes[i].No);
                    modeshape.AddMode(j + 1, result.StdDisplacementX, result.StdDisplacementY, result.StdDisplacementZ);
                }
                modeShapes.Add(modeshape);
            }
            return modeShapes;
        }

        /// <summary>
        /// Get all the standardized displacement for the different mode shapes and FE nodes
        /// </summary>
        /// <param name="model">RFEM Model</param>
        /// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
        /// <returns></returns>
        public static List<ModeShape> GetAllStandardizedDisplacement(IModel model, int NVC)
        {

            var DynamPro = model.GetModule("DynamPro") as IDynamModule;

            IDynamResults dynamicResults = DynamPro.GetResults();

            int nbModeShape = DynamPro.GetData().GetNVCParams(NVC).NumberModeShapes;
            var MeshNodes = model.GetCalculation().GetFeMesh().GetNodes();
            int nbMeshNode = MeshNodes.Length;

            var shapes = dynamicResults.GetAllNVCModeShapes(NVC);

            List<ModeShape> modeShapes = new List<ModeShape>();

            for (int i = 0; i < nbMeshNode; i++)
            {
                ModeShape modeshape = new ModeShape(MeshNodes[i].No);
                modeshape.Modes.Add(shapes[i].ModeShape);
                modeshape.Ux.Add(shapes[i].StdDisplacementX);
                modeshape.Uy.Add(shapes[i].StdDisplacementY);
                modeshape.Uz.Add(shapes[i].StdDisplacementZ);

                modeShapes.Add(modeshape);
            }

            for (int i = 0; i < nbModeShape - 1; i++)
            {
                for (int j = 0; j < nbMeshNode; j++)
                {
                    modeShapes[j].Modes.Add(shapes[nbMeshNode * (i + 1) + j].ModeShape);
                    modeShapes[j].Ux.Add(shapes[nbMeshNode * (i + 1) + j].StdDisplacementX);
                    modeShapes[j].Uy.Add(shapes[nbMeshNode * (i + 1) + j].StdDisplacementY);
                    modeShapes[j].Uz.Add(shapes[nbMeshNode * (i + 1) + j].StdDisplacementZ);
                }

            }

            return modeShapes;
        }


        public static void GetDymanicData(IModel model, int NVC)
        {

            var DynamPro = model.GetModule("DynamPro") as IDynamModule;
            IDynamResults dynamicResults = DynamPro.GetResults();

            var naturalFrequencies = dynamicResults.GetAllNVCNaturalFrequencies(1);
            var modalMasses = dynamicResults.GetAllNVCMassFactors(1);
            var shapes = dynamicResults.GetAllNVCModeShapes(1);
            var shapesFEPoint = dynamicResults.GetNVCModeShapesOnFEPoint(1, 1, 1);


            IModelData ModelData = model.GetModelData();
            ISurface srf = ModelData.GetSurface(1, ItemAt.AtNo);

            //Get FE nodes
            ICalculation calcs = model.GetCalculation();
            IFeMesh mesh = calcs.GetFeMesh();
            var FEnodes = mesh.GetNodes();
            var Mesh2D = mesh.Get2DElements();

        }

        /// <summary>
        /// Return a list of all FE nodes in the model
        /// </summary>
        /// <param name="model">RFEM Model</param>
        /// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
        /// <returns></returns>
        public static List<FeMeshNode> GetFENodes(IModel model, int NVC)
        {
           return model.GetCalculation().GetFeMesh().GetNodes().ToList();
        }






    }
}
