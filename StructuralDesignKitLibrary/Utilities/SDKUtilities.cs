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

    }
}
