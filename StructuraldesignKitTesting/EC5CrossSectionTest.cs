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
using System.IO;
using System;
using StructuralDesignKitLibrary.CrossSections.Interfaces;

namespace StructuraldesignKitTesting
{
    public class EC5CrossSectionTest
    {

        private static string TestFilePath = "C:\\Users\\Guillaume Caussarieu\\source\\repos\\StructuralDesignKit_Holz\\StructuraldesignKitTesting\\TestData\\Test_CrossSections.xlsx";
        private static Excel.Application XlApp = new Excel.Application();




        [Theory]
        [MemberData(nameof(GetTensionParallelToGrainData))]
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



        [Theory]
        [MemberData(nameof(GetCompressionParallelToGrainData))]
        public void CompressionParallelToGrainTest(double Sig0_t_d, IMaterial material, double Kmod, double Ym, double expected)
        {
            //Arrange
            //data generated previsously


            //Act
            double actual = StructuralDesignKitLibrary.EC5.EC5_CrossSectionCheck.CompressionParallelToGrain(Sig0_t_d, material, Kmod, Ym);

            //Assert
            //Tolerance is the digit looked at. i.e : 0.0X<- digit compared
            Assert.Equal(expected, actual, 2);
        }

        [Theory]
        [MemberData(nameof(GetBending_6_1_6Data))]
        public void Bending_6_1_6Test(double SigMyd, double SigMzd, ICrossSection crossSection, IMaterial material, double Kmod, double Ym, double khy, double khz, double expected)
		{
			//Arrange
			//data generated previsously

			//Act
			double actual = StructuralDesignKitLibrary.EC5.EC5_CrossSectionCheck.Bending(SigMyd, SigMzd, crossSection, material, Kmod, Ym, khy, khz);

			//Assert
			//Tolerance is the digit looked at. i.e : 0.0X<- digit compared
			Assert.Equal(expected, actual, 2);
		}




		



        public static IEnumerable<object[]> GetTensionParallelToGrainData()
        {

            Workbook wb = XlApp.Workbooks.Open(TestFilePath);

            var ws = GetDataFromExcelTab("TensionParallelToGrain");

            Range cell = ws.Cells[3, 1];
            while (cell.Value2 != null)
            {
                int b = (int)cell.Value2;
                int h = (int)cell.Offset[0, 1].Value2;
                IMaterialTimber mat = StructuralDesignKitExcel.ExcelHelpers.GetTimberMaterialFromTag(cell.Offset[0, 2].Value2);
                var CS = new StructuralDesignKitLibrary.CrossSections.CrossSectionRectangular(b, h, mat);
                double stress = CS.ComputeNormalStress(cell.Offset[0, 3].Value2);
                double kh = EC5_Factors.Kh_Tension(mat.Type, CS.B, CS.H);
                double kmod = (double)cell.Offset[0, 4].Value2;
                double Ym = (double)cell.Offset[0, 5].Value2;
                double ratio = (double)cell.Offset[0, 6].Value2;
                cell = cell.Offset[1, 0];

                yield return new object[] { stress, mat, kmod, Ym, kh, 1, ratio };
            }
            XlApp.Workbooks.Close();
            XlApp.Quit();
        }


        public static IEnumerable<object[]> GetCompressionParallelToGrainData()
        {

            Workbook wb = XlApp.Workbooks.Open(TestFilePath);

            var ws = GetDataFromExcelTab("CompressionParallelToGrain");

            Range cell = ws.Cells[3, 1];
            while (cell.Value2 != null)
            {
                int b = (int)cell.Value2;
                int h = (int)cell.Offset[0, 1].Value2;
                IMaterialTimber mat = StructuralDesignKitExcel.ExcelHelpers.GetTimberMaterialFromTag(cell.Offset[0, 2].Value2);
                var CS = new StructuralDesignKitLibrary.CrossSections.CrossSectionRectangular(b, h, mat);
                double stress = CS.ComputeNormalStress(cell.Offset[0, 3].Value2);
                double kmod = (double)cell.Offset[0, 4].Value2;
                double Ym = (double)cell.Offset[0, 5].Value2;
                double ratio = (double)cell.Offset[0, 6].Value2;
                cell = cell.Offset[1, 0];

                yield return new object[] { stress, mat, kmod, Ym,ratio };
            }
            XlApp.Workbooks.Close();
            XlApp.Quit();
        }

		public static IEnumerable<object[]> GetBending_6_1_6Data()
		{

			Workbook wb = XlApp.Workbooks.Open(TestFilePath);

			var ws = GetDataFromExcelTab("Bending_6.1.6");

			Range cell = ws.Cells[3, 1];
			while (cell.Value2 != null)
			{
				int b = (int)cell.Value2;
				int h = (int)cell.Offset[0, 1].Value2;
				IMaterialTimber mat = StructuralDesignKitExcel.ExcelHelpers.GetTimberMaterialFromTag(cell.Offset[0, 2].Value2);
				var CS = new StructuralDesignKitLibrary.CrossSections.CrossSectionRectangular(b, h, mat);
				double stressMy = CS.ComputeStressBendingY(cell.Offset[0, 3].Value2);
				double stressMz = CS.ComputeStressBendingZ(cell.Offset[0, 4].Value2);
				double kmod = (double)cell.Offset[0, 5].Value2;
				double Ym = (double)cell.Offset[0, 6].Value2;
				double ratio = (double)cell.Offset[0, 7].Value2;
                double khy = EC5_Factors.Kh_Bending(mat.Type, CS.H);
                double khz = 1;
				// For verification with RFEM 5, which does not take into account the kh factor for the z axis for Glulam
				if (mat.Type == EC5_Utilities.TimberType.Softwood || mat.Type == EC5_Utilities.TimberType.Hardwood)
				{
                    khz = EC5_Factors.Kh_Bending(mat.Type, CS.B);
				}

				cell = cell.Offset[1, 0];

				yield return new object[] { stressMy, stressMz,CS, mat, kmod, Ym,khy,khz, ratio };
			}
			XlApp.Workbooks.Close();
			XlApp.Quit();
		}


		#region Utilities
		/// <summary>
		/// return a worksheet based on the tab name
		/// </summary>
		/// <param name="tabName"></param>
		/// <returns></returns>
		private static Worksheet GetDataFromExcelTab(string tabName)
        {
            List<Worksheet> ws = new List<Worksheet>();

            foreach (Worksheet sheet in XlApp.Worksheets)
            {
                ws.Add(sheet);
            }

            return ws.Where(p => p.Name == tabName).First();
        }
        #endregion
    }
}
