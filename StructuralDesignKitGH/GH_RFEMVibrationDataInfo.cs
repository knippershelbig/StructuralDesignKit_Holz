using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace StructuralDesignKitGH
{
    public class GH_RFEMVibrationDataInfo : GH_AssemblyInfo
    {
        public override string Name => "StructuralDesignKitGH";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("760C0F72-99A7-496A-8840-D7CDB544087A");

        //Return a string identifying you or your company.
        public override string AuthorName => "";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";
    }
}