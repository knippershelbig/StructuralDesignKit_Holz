using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using StructuralDesignKitLibrary;
using StructuralDesignKitLibrary.RFEM;
using StructuralDesignKitGrasshopper;
using Dlubal.RFEM5;

namespace StructuralDesignKitGH
{
    public class GH_RFEMVibrationData : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public GH_RFEMVibrationData()
          : base("Get RFEM Vibration data", "VibrationData",
            "Get the data necessary to perform a vibration analysis from an open RFEM model",
            "SDK", "Vibrations")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Run", "Run", "Get the data from the active RFEM model using the COM interface", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data Vibration", "data", "RFEM data necessary for vibrationpost processing", GH_ParamAccess.item);
        }


        //Persistant data for the solve instance method
        RFEMVibrationDataObject data = new RFEMVibrationDataObject();

        
        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            bool run = false;


            DA.GetData(0, ref run);

            if (run)
            {
                var RFEM_Utilities = new RFEM5_Utilities();

                var model = RFEM_Utilities.OpenModel();

                var FEnodes = RFEM_Utilities.GetFENodes(model, 1);
                var naturalFrequencies = RFEM_Utilities.GetNaturalFrequencies(model, 1);
                var modalMasses = RFEM_Utilities.GetModalMasses(model, 1);
                var modeDisplacement = RFEM_Utilities.GetAllStandardizedDisplacement(model, 1);

                List<int> modes = new List<int>();
                for (int i = 0; i < naturalFrequencies.Count; i++)
                {
                    modes.Add(i + 1);
                }

                data = new RFEMVibrationDataObject(modes, FEnodes, naturalFrequencies, modalMasses, modeDisplacement);


                RFEM_Utilities.CloseRFEMModel(model);
                
            }
                DA.SetData(0, data);

        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return StructuralDesignKitGH.Properties.Resources.IconVibrationData;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("752D2FB3-7894-47D2-B806-8132EF130B13");
    }
}