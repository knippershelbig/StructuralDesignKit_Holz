using System;
using System.Collections.Generic;
using System.Diagnostics;
using Grasshopper.Kernel;
using Rhino.FileIO;
using Rhino.Geometry;
using StructuralDesignKitGrasshopper;
using StructuralDesignKitLibrary.Vibrations;

namespace StructuralDesignKitGH
{
    public class GH_RFEM_Resonant_ComputeTwoPoints : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_RFEM_Resonant_ComputeTwoPoints class.
        /// </summary>
        public GH_RFEM_Resonant_ComputeTwoPoints()
          : base("Vibration Resonant response 2 Points", "ResonantResponse 2 Points",
              "Compute the resonant response from a structure for an excitation and response point",
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
            pManager.AddNumberParameter("walkingLength  ", "WL", "Length of the walking path, if negative, the Eurocode resonant build up factor is considered", GH_ParamAccess.item, -1);
            pManager.AddPointParameter("Excitation Point", "exPt", "Point where the excitation is produced", GH_ParamAccess.item);
            pManager.AddPointParameter("Response Points", "RespPt", "Points where the excitation is percieved", GH_ParamAccess.list);
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

            Point3d ExciPoint = new Point3d();
            List<Point3d> RespPointList = new List<Point3d>();


            DA.GetData(0, ref data);
            DA.GetData(1, ref importFc);
            DA.GetData(2, ref Xi);
            DA.GetData(3, ref weighting);
            DA.GetData(4, ref responseFactor);
            DA.GetData(5, ref WalkLen);
            DA.GetData(6, ref ExciPoint);
            DA.GetDataList(7, RespPointList);


            List<double> responses = new List<double>();
            List<Point3d> ResultPoints = new List<Point3d>();



            string typeValue = importFc.GetType().ToString();

            if (importFc.GetType().ToString() == "Grasshopper.Kernel.Types.GH_String")
            {
                fp = double.Parse(importFc.ToString());
            }
            else if (importFc.GetType() == typeof(double)) fp = (double)importFc;
            fp = double.Parse(importFc.ToString());

            //find closest FE node from excitation and response points
            //Rtree to implement


            List<Point3d> FENodes = new List<Point3d>();

            foreach (var pt in data.FENodes)
            {
                FENodes.Add(new Point3d(pt.X, pt.Y, pt.Z));
            }

            //find closest FE node from the excitation point:
            int closestExcitation = 0;
            double minDistExcitation = 1e6;
            int count = 0;

            foreach (Point3d node in FENodes)
            {
                if (ExciPoint.DistanceTo(node) < minDistExcitation)
                {
                    closestExcitation = count;
                    minDistExcitation = ExciPoint.DistanceTo(node);
                }
                count += 1; ;
            }
            Point3d ExPoint = FENodes[closestExcitation];


            foreach (Point3d pt in RespPointList)
            {
                int closestResponse =0;
                double minDistResponse = 1e6;

                count = 0;
                foreach (Point3d node in FENodes)
                {
                    if (pt.DistanceTo(node) < minDistResponse)
                    {
                        closestResponse = count;
                        minDistResponse = pt.DistanceTo(node);
                    }
                    count += 1; ;
                }

                Vibrations.Weighting W = Vibrations.GetWeighting(weighting);

                responses.Add(Vibrations.ResonantResponseAnalysis(data.ModeShapes[closestExcitation].Uz, data.ModeShapes[closestResponse].Uz, data.NaturalFrequencies, data.ModalMasses, fp, Xi, W, WalkLen, responseFactor));
                ResultPoints.Add(FENodes[closestResponse]);
                
            }


            DA.SetDataList(0, ResultPoints);
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
            get { return new Guid("1B5F1EEF-9731-481E-927B-F86A76A3A653"); }
        }
    }
}