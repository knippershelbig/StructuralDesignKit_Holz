using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructuralDesignKitExcel
{
    public static partial class ExcelFormulae
    {
        #region utilities
        [ExcelFunction(Description = "Create a cross section tag",
            Name = "SDK.Utilities.CreateRectangularCrossSection",
            IsHidden = false,
            Category = "SDK.Utilities")]
        public static string CreateCrossSection([ExcelArgument(Description = "width")] double b, [ExcelArgument(Description = "height")] double h, string material)
        {
            return ExcelHelpers.CreateRectangularCrossSectionTag(b, h, ExcelHelpers.GetTimberMaterialFromTag(material));
        }


        [ExcelFunction(Description = "Create a Bolt tag",
            Name = "SDK.Utilities.CreateBoltTag",
            IsHidden = false,
            Category = "SDK.Utilities")]
        public static string CreateBoltTag(
            [ExcelArgument(Description = "Diameter of the fastener")] double diameter,
            [ExcelArgument(Description = "Tensile strength of the fasterner in N/mm²")] double fu)
        {
            return ExcelHelpers.GenerateBoltTag(diameter, fu);

        }

        [ExcelFunction(Description = "Create a Dowel tag",
            Name = "SDK.Utilities.CreateDowelTag",
            IsHidden = false,
            Category = "SDK.Utilities")]
        public static string CreateDowelTag(
            [ExcelArgument(Description = "Diameter of the fastener")] double diameter,
            [ExcelArgument(Description = "Tensile strength of the fasterner in N/mm²")] double fu)
        {
            return ExcelHelpers.GenerateDowelTag(diameter, fu);

        }
        #endregion
    }
}
