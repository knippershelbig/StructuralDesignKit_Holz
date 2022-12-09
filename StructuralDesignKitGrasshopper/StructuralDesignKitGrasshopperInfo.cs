using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace StructuralDesignKitGrasshopper
{
    public class StructuralDesignKitGrasshopperInfo : GH_AssemblyInfo
    {
        public override string Name => "StructuralDesignKitGrasshopper";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("0B7682CD-2F26-4FA8-832A-DC480E393C35");

        //Return a string identifying you or your company.
        public override string AuthorName => "";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";
    }
}