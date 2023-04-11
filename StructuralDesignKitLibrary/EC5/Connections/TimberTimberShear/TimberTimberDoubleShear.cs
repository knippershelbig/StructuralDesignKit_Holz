using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Materials;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Connections.TimberTimberShear
{
    public class TimberTimberDoubleShear : ITimberTimberShear
    {

        #region Properties
        public IFastener Fastener { get; set; }
        public double Angle1 { get; set; }
        public double Angle2 { get; set; }
        public IMaterialTimber Timber1 { get; set; }
        public IMaterialTimber Timber2 { get; set; }
        public double T1 { get; set; }
        public double T2 { get; set; }

        public List<string> FailureModes { get; set; }
        public List<double> Capacities { get; set; }
        public string FailureMode { get; set; }

        /// <summary>
        /// Characteristic strength for 2 shear planes
        /// </summary>
        public double Capacity { get; set; }
        public bool RopeEffect { get; set; }

        public double Fhk1 { get; set; }
        public double Fhk2 { get; set; }

        #endregion

        public TimberTimberDoubleShear(IFastener fastener, IMaterialTimber timber1, double timberThickness1, double angle1, IMaterialTimber timber2, double timberThickness2, double angle2, bool ropeEffect)
        {
            Fastener = fastener;
            Angle1 = angle1;
            Angle2 = angle2;
            Timber1 = timber1;
            Timber2 = timber2;
            T1 = timberThickness1;
            T2 = timberThickness2;
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
            //Embedment strength timber 1
            Fastener.ComputeEmbedmentStrength(Timber1, Angle1);
            Fhk1 = Fastener.Fhk;

            //Embedment strength timber 2
            Fastener.ComputeEmbedmentStrength(Timber2, Angle2);
            Fhk2 = Fastener.Fhk;


            double capacity = 0;
            double RopeEffectCapacity = 0;
            double B = Fhk2 / Fhk1;

            //Failure mode according to EN 1995-1-1 Eq (8.7)

            //Failure mode g
            FailureModes.Add("g");
            Capacities.Add(Fhk1 * T1 * Fastener.Diameter);


            //Failure mode h
            FailureModes.Add("h");
            Capacities.Add(0.5*Fhk2 * T2 * Fastener.Diameter);


            //Failure mode j
            FailureModes.Add("j");
            capacity = 1.05 * Capacities[0] / (2 + B) * (Math.Sqrt(2 * B * (1 + B) + 4 * B * (2 + B) * Fastener.MyRk / (Fhk1 * Fastener.Diameter * Math.Pow(T1, 2))) - B);
            if (RopeEffect)
            {
                Fastener.ComputeWithdrawalStrength(this);
                RopeEffectCapacity = Fastener.WithdrawalStrength / 4;
                capacity += Math.Min(Fastener.MaxJohansenPart * capacity, RopeEffectCapacity);
            }
            Capacities.Add(capacity);


            //Failure mode k
            FailureModes.Add("k");
            capacity = 1.15 * Math.Sqrt(2 * B / (1 + B)) * Math.Sqrt(2 * Fastener.MyRk * Fhk1 * Fastener.Diameter);
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
