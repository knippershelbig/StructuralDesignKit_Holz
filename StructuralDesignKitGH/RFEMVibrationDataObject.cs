using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dlubal.RFEM5;
using Dlubal.DynamPro;
using StructuralDesignKitLibrary.RFEM;

namespace StructuralDesignKitGrasshopper
{
    internal class RFEMVibrationDataObject
    {

        //list of modes
        //list of nodes
        //List of natural frequencies
        //List of modal masses
        //List of standarized displacement

        public List<int> Modes{ get; set; }
        public List<FeMeshNode> FENodes{ get; set; }
        public List<double> ModalMasses{ get; set; }
        public List<double> NaturalFrequencies{ get; set; }
        public List<ModeShape> ModeShapes { get; set; }

        public RFEMVibrationDataObject()
        {

        }
        public RFEMVibrationDataObject(List<int> modes, List<FeMeshNode> feNodes,List<double> naturalFrequencies, List<double> modalMasses, List<ModeShape> modeShapes  )
        {
            Modes = modes;
            NaturalFrequencies = naturalFrequencies;
            FENodes = feNodes;
            ModalMasses = modalMasses;
            ModeShapes = modeShapes; 
        }
    }
}
