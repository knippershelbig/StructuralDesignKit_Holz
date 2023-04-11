using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.SteelTimberShear
{
    public class SingleOuterSteelPlate : ISteelTimberShear
    {


        public IFastener Fastener { get; set; }
        public double SteelPlateThickness { get; set; }
        public bool isTkickPlate { get; set; }
        public bool isThinPlate { get; set; }
        public double Angle { get; set; }
        public IMaterialTimber Timber { get; set; }
        public double TimberThickness { get; set; }
        public List<string> FailureModes { get; set; }
        public List<double> Capacities { get; set; }
        public string FailureMode { get; set; }
        public double Capacity { get; set; }
        public bool RopeEffect { get; set; }



        public SingleOuterSteelPlate(IFastener fastener, double steelPlateThickness, double angle, IMaterialTimber timber, double timberThickness, bool ropeEffect)
        {
            Fastener = fastener;
            SteelPlateThickness = steelPlateThickness;
            Angle = angle;
            Timber = timber;
            TimberThickness = timberThickness;
            RopeEffect = ropeEffect;

            //Initialize lists
            FailureModes = new List<string>();
            Capacities = new List<double>();

            ComputeFailingModes();

            double capacityThinPlate = Math.Min(Capacities[0], Capacities[1]);
            double capacityThickPlate = Capacities.GetRange(2, 3).Min();

            //case thin plate
            if (SteelPlateThickness <= 0.5 * Fastener.Diameter) Capacity = capacityThinPlate;

            //case thick plate
            else if (SteelPlateThickness >= Fastener.Diameter) Capacity = capacityThickPlate;

            //Case interpolation between thin and thick plate
            else Capacity = Utilities.SDKUtilities.LinearInterpolation(steelPlateThickness, 0.5 * Fastener.Diameter, capacityThinPlate, Fastener.Diameter, capacityThickPlate);

            FailureMode = FailureModes[Capacities.IndexOf(Capacities.Min())];
        }


        public void ComputeFailingModes()
        {
            Fastener.ComputeEmbedmentStrength(Timber, Angle);
            double capacity = 0;
            Fastener.ComputeWithdrawalStrength(this);
            double RopeEffectCapacity = Fastener.WithdrawalStrength / 4;

            //Failure mode according to EN 1995-1-1 Eq (8.10)

            //thin steel plate
            //Failure mode a
            FailureModes.Add("a");
            Capacities.Add(0.4 * Fastener.Fhk * TimberThickness * Fastener.Diameter);

            //Failure mode b
            FailureModes.Add("b");
            capacity = 1.15 * Math.Sqrt(2 * Fastener.MyRk * Fastener.Fhk * Fastener.Diameter);
            if (RopeEffect) capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            Capacities.Add(capacity);

            //thick steel plate
            //Failure mode 
            FailureModes.Add("c");
            Capacities.Add(Fastener.Fhk * TimberThickness * Fastener.Diameter);

            FailureModes.Add("d");
            capacity = Capacities[2] * (Math.Sqrt(2 + 4 * Fastener.MyRk / (Fastener.Fhk * Fastener.Diameter * Math.Pow(TimberThickness, 2))) - 1);
            if (RopeEffect) capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            Capacities.Add(capacity);

            FailureModes.Add("e");
            capacity = 2.3 * Math.Sqrt(Fastener.MyRk * Fastener.Fhk * Fastener.Diameter);
            if (RopeEffect) capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            Capacities.Add(capacity);
        }
    }
}
