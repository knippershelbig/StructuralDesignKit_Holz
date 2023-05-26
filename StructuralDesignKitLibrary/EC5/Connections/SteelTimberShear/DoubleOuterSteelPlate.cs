using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.SteelTimberShear
{
    public class DoubleOuterSteelPlate : ISteelTimberShear
    {

        #region properties
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

        /// <summary>
        /// Characteristic strength for 2 shear planes
        /// </summary>
        public double Capacity { get; set; }
        public bool RopeEffect { get; set; }
        #endregion


        public DoubleOuterSteelPlate(IFastener fastener, double steelPlateThickness, double angle, IMaterialTimber timber, double timberThickness, bool ropeEffect)
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
            double capacityThickPlate = Capacities.GetRange(2, 2).Min();

            //case thin plate
            if (SteelPlateThickness <= 0.5 * Fastener.Diameter) Capacity = capacityThinPlate;

            //case thick plate
            else if (SteelPlateThickness >= Fastener.Diameter) Capacity = capacityThickPlate;

            //Case interpolation between thin and thick plate
            else Capacity = Utilities.SDKUtilities.LinearInterpolation(steelPlateThickness, 0.5 * Fastener.Diameter, capacityThinPlate, Fastener.Diameter, capacityThickPlate);

            FailureMode = FailureModes[Capacities.IndexOf(Capacities.Min())];

            Capacity = Capacity * 2;

            FailureMode = FailureModes[Capacities.IndexOf(Capacities.Min())];
        }


        public void ComputeFailingModes()
        {
            Fastener.ComputeEmbedmentStrength(Timber, Angle);
            double capacity = 0;
            double RopeEffectCapacity = 0;


            //Failure mode according to EN 1995-1-1 Eq (8.12 and 8.13)

            //thin steel plates
            //Failure mode j
            FailureModes.Add("j");
            Capacities.Add(0.5 * Fastener.Fhk * TimberThickness * Fastener.Diameter);

            //Failure mode k
            FailureModes.Add("k");
            capacity = 1.15 * Math.Sqrt(2 * Fastener.MyRk * Fastener.Fhk * Fastener.Diameter);
            if (RopeEffect)
            {
                Fastener.ComputeWithdrawalStrength(this);
                RopeEffectCapacity = Fastener.WithdrawalStrength / 4;
                capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            }
            Capacities.Add(capacity);


            //thick steel plates

            //Failure mode l
            FailureModes.Add("l");
            Capacities.Add(0.5 * Fastener.Fhk * TimberThickness * Fastener.Diameter);

            //Failure mode m
            FailureModes.Add("m");
            capacity = 2.3 * Math.Sqrt(Fastener.MyRk * Fastener.Fhk * Fastener.Diameter);
            if (RopeEffect)
            {
                Fastener.ComputeWithdrawalStrength(this);
                RopeEffectCapacity = Fastener.WithdrawalStrength / 4;
                capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            }
            Capacities.Add(capacity);
        }
    }
}
