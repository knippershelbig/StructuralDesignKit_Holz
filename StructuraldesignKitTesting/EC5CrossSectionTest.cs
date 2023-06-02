using Microsoft.Office.Interop.Excel;
using StructuralDesignKitLibrary.Materials;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using StructuralDesignKitLibrary;
using StructuralDesignKitLibrary.EC5;
using System.Globalization;
using System.Reflection;
using Dlubal.WS.Rfem6.Model;

namespace StructuraldesignKitTesting
{
    public class EC5CrossSectionTest
    {


        //private static string Path = @"U:\Desktop\tests\Test_CrossSections.xlsx";
        private static Excel.Application XlApp = new Excel.Application();



        private static Worksheet PopulateTests(string path)
        {

            //Read a Excel file, with one tab per check type
            Workbook wb = XlApp.Workbooks.Open(path);
            List<Worksheet> ws = new List<Worksheet>();
            foreach (Worksheet sheet in XlApp.Worksheets)
            {
                ws.Add(sheet);
            }

            var tensionWS = ws.Where(p => p.Name == "TensionParallelToGrain").First();

            return tensionWS;
        }

        private static void CloseExcelApp()
        {
            XlApp.Workbooks.Close();
            XlApp.Quit();
        }


        public static IEnumerable<object[]> GetTensionData()
        {
            var ws = PopulateTests(GetTestDataPath());


            Range cell = ws.Cells[3, 1];
            while (cell.Value2 != null)
            {
                int b = (int)cell.Value2;
                int h = (int)cell.Offset[0, 1].Value2;
                IMaterialTimber mat = StructuralDesignKitExcel.ExcelHelpers.GetTimberMaterialFromTag(cell.Offset[0, 2].Value2);
                var CS = new StructuralDesignKitLibrary.CrossSections.CrossSectionRectangular(b, h, mat);
                double stress = CS.ComputeNormalStress(cell.Offset[0, 3].Value2);
                double kh = EC5_Factors.Kh_Tension(mat.Type, b);
                double kmod = (double)cell.Offset[0, 4].Value2;
                double Ym = (double)cell.Offset[0, 5].Value2;
                double ratio = (double)cell.Offset[0, 6].Value2;
                cell = cell.Offset[1, 0];

                yield return new object[] { stress, mat, kmod, Ym, kh, 1, ratio };
            }
            CloseExcelApp();
        }



        /// <summary>
        /// return the path of the excel file which contains the test data
        /// </summary>
        private static string GetTestDataPath()
        {
           // String fileName = "..\\TestData\\TestCrossSection.xlsx";

           string fileName = "TestCrossSection.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            return fileName;
        }





        [Theory]
        [MemberData(nameof(GetTensionData))]
        public void TensionParallelToGrainTest(double Sig0_t_d, IMaterial material, double Kmod, double Ym, double Kh, double Kl_LVL, double expected)
        {
            //Arrange
            //data generated previsously


            //Act
            double actual = StructuralDesignKitLibrary.EC5.EC5_CrossSectionCheck.TensionParallelToGrain(Sig0_t_d, material, Kmod, Ym, Kh, Kl_LVL);

            //Assert
            //Tolerance is the digit looked at. i.e : 0.0X<- digit compared
            Assert.Equal(expected, actual, 2);
        }
    }
}