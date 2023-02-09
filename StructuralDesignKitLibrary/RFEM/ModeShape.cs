using System.Collections.Generic;

namespace StructuralDesignKitLibrary.RFEM
{
    /// <summary>
    /// class representing the mode shape standadized displacements in x, y, z
    /// </summary>
    public class ModeShape
    {
        public int NodeID { get; set; }
        public List<int> Modes { get; set; }
        public List<double> Ux { get; set; }
        public List<double> Uy { get; set; }
        public List<double> Uz { get; set; }


        public ModeShape(int id)
        {
            NodeID = id;
            Modes = new List<int>();
            Ux = new List<double>();
            Uy = new List<double>();
            Uz = new List<double>();
        }

        public ModeShape(int id, List<int> modes, List<double> ux, List<double> uy, List<double> uz)
        {
            NodeID = id;
            Modes = modes;
            Ux= ux;
            Uy = uy;
            Uz= uz;     
        }

        public void AddMode(int mode, double ux, double uy, double uz)
        {
            Modes.Add(mode);
            Ux.Add(ux);
            Uy.Add(uy);
            Uz.Add(uz);
        }
    }
}
