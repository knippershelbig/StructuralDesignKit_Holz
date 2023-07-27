using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.Connections.Interface;
using StructuralDesignKitLibrary.Connections.SteelTimberShear;
using StructuralDesignKitLibrary.Connections.TimberTimberShear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static StructuralDesignKitLibrary.EC5.EC5_Utilities;
using Microsoft.Office.Interop.Excel;
using System.Windows.Controls;
using StructuralDesignKitExcel;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;

namespace StructuralDesignKitExcel.RibbonActions
{
    public class FireButtonActions
    {
        public static void ValidateCellWithPlasterboardTypes(Excel.Application xlApp)
        {

            var activeCell = xlApp.ActiveCell;
            var plasterboards = StructuralDesignKitExcel.ExcelHelpers.GetPlasterboardTypes();
            plasterboards.Add("none");

            RibbonActions.RibbonUtilities.ValidateCellWithList(activeCell, plasterboards);




        }
    }


}

