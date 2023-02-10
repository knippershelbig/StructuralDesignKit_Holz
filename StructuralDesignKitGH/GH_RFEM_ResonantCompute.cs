using System;
using System.Collections.Generic;
using Dlubal.DynamPro;
using Grasshopper.Kernel;
using Rhino.Geometry;
using StructuralDesignKitGrasshopper;
using StructuralDesignKitLibrary.Vibrations;

namespace StructuralDesignKitGH
{
    public class GH_RFEM_ResonantCompute : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GH_RFEM_ResonantCompute()
          : base("Vibration Resonant response", "ResonantResponse",
              "Compute the resonant response from a structure",
              "SDK", "Vibration")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            string weighting = "Workshop\r\nCirculationSpace\r\nResidential\r\nOffice\r\n Ward\r\nGeneralLaboratory\r\nConsultingRoom\r\nCriticalWorkingArea\r\nNone";
            pManager.AddGenericParameter("data", "data", "RFEM vibration data from the component VibrationData", GH_ParamAccess.item);
            pManager.AddGenericParameter("PaceFrequency", "fp", "Pace frequency, either as single value or as range", GH_ParamAccess.item);
            pManager.AddNumberParameter("Damping ratio", "Xi", "Damping ratio", GH_ParamAccess.item);
            pManager.AddTextParameter("weigthingCategory", "W", weighting, GH_ParamAccess.item);
            pManager.AddBooleanParameter("Response Factor", "R", "If true, provide the Response factor instead of the acceleration", GH_ParamAccess.item);
            pManager.AddNumberParameter("walkingLength  ", "WL", "Length of the walking path, if negative, the Eurocode resonant build up factor is considered", GH_ParamAccess.item,-1);
            pManager[5].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddPointParameter("points", "pt", "", GH_ParamAccess.list);
            pManager.AddNumberParameter("Response", "R", "Response value", GH_ParamAccess.list);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            RFEMVibrationDataObject data = null;
            object importFc = null;
            double fp = 0;
            double Xi = 0;
            string weighting = "";
            bool responseFactor = false;
            double WalkLen = -1;

            DA.GetData(0, ref data);
            DA.GetData(1, ref importFc);
            DA.GetData(2, ref Xi);
            DA.GetData(3, ref weighting);
            DA.GetData(4, ref responseFactor);
            DA.GetData(5, ref WalkLen);

            string typeValue = importFc.GetType().ToString();

            if (importFc.GetType().ToString() == "Grasshopper.Kernel.Types.GH_String")
            {
                fp = double.Parse(importFc.ToString());
            }
            else if(importFc.GetType() == typeof(double))fp = (double)importFc ;
            fp = double.Parse(importFc.ToString());
            List<Point3d> pts = new List<Point3d>();
            foreach (var pt in data.FENodes)
            {
                pts.Add(new Point3d(pt.X, pt.Y, pt.Z));
            }

            List<double> responses = new List<double>();

            Vibrations.Weighting W = Vibrations.GetWeighting(weighting);

            foreach (var item in data.ModeShapes)
            {

                responses.Add(Vibrations.ResonantResponseAnalysis(item.Uz, item.Uz, data.NaturalFrequencies, data.ModalMasses, fp, Xi, W, WalkLen, responseFactor));
            }

            DA.SetDataList(0, pts);
            DA.SetDataList(1, responses);


        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("8D214E3D-E757-4699-B95D-1DBFDF715789"); }
        }
    }
}