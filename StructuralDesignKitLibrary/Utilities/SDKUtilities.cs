using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitLibrary.Utilities
{
    public static class SDKUtilities
    {

        public static double LinearInterpolation(double x, double x0, double y0, double x1, double y1)
        {
            return (y0 * (x1 - x) + y1 * (x - x0)) / (x1 - x0);
        }


        public static int ComputeAngleToFirstQuadrant(double angle)
        {
            int angleRounded = Math.Abs(Convert.ToInt32(angle));
            int angleRemainder;
            int angleFirstQuadrant;

            if (angleRounded <= 90) return angleRounded;

            else
            {
                Math.DivRem(angleRounded, 360, out angleRemainder);

                if (angleRemainder <= 90) angleFirstQuadrant = angleRemainder;

                else if (angleRemainder > 90 && angleRemainder <= 180) angleFirstQuadrant = 180 - angleRemainder;

                else if (angleRemainder > 180 && angleRemainder <= 270) angleFirstQuadrant = angleRemainder - 180;

                else angleFirstQuadrant = 360 - angleRemainder;

                return angleFirstQuadrant;

            }
        }

    }

}
