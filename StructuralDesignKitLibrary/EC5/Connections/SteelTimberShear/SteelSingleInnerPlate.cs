using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.SteelTimberShear
{
    public class SteelSingleInnerPlate : ISteelTimberShear
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
        


        public SteelSingleInnerPlate(IFastener fastener, double steelPlateThickness, double angle, IMaterialTimber timber, double timberThickness, bool ropeEffect)
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

            Capacity = Capacities.Min() * 2;
            FailureMode = FailureModes[Capacities.IndexOf(Capacities.Min())];



        }



        public void ComputeFailingModes()
        {
            Fastener.ComputeEmbedmentStrength(Timber, Angle);
            double capacity = 0;
            double RopeEffectCapacity = 0;


            //Failure mode according to EN 1995-1-1 Eq (8.10)

            //Failure mode f
            FailureModes.Add("f");
            Capacities.Add(Fastener.Fhk * TimberThickness * Fastener.Diameter);

            //Failure mode g
            FailureModes.Add("g");
            capacity = Capacities[0] * (Math.Sqrt(2 + 4 * Fastener.MyRk / (Fastener.Fhk * Fastener.Diameter * Math.Pow(TimberThickness, 2))) - 1);
            if (RopeEffect)
            {
                Fastener.ComputeWithdrawalStrength(this);
                RopeEffectCapacity = Fastener.WithdrawalStrength / 4;
                capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            }
            Capacities.Add(capacity);

            //Failure mode h
            FailureModes.Add("h");
            capacity = 2.3 * Math.Sqrt(Fastener.MyRk * Fastener.Fhk * Fastener.Diameter);
            if (RopeEffect) capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            Capacities.Add(capacity);

        }
    }
}
