using StructuralDesignKitLibrary.Materials;
using Xunit;

namespace StructuraldesignKitTesting
{
    public class EC5CrossSectionTest
    {



        public static IEnumerable<object[]> Tension =>
            new List<object[]>
            {
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h),0.6,1.3,1,1, 0.113 },
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL28h),0.6,1.3,1,1, 0.097 },
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h),0.8,1.3,1,1, 0.085 },
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL28h),0.8,1.3,1,1, 0.073 },
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL24h),1,1.3,1,1, 0.068},
                new object[]{1, new MaterialTimberGlulam(MaterialTimberGlulam.Grades.GL28h),1,1.3,1,1, 0.058 }
            };



        [Theory]
        [MemberData(nameof(Tension))]
        public void TensionParallelToGrainTest(double Sig0_t_d, IMaterial material, double Kmod, double Ym, double Kh, double Kl_LVL, double expected)
        {
            //Arrange
            //data generated previsously


            //Act
            double actual = StructuralDesignKitLibrary.EC5.EC5_CrossSectionCheck.TensionParallelToGrain(Sig0_t_d, material, Kmod, Ym, Kh, Kl_LVL);

            //Assert
            //Tolerance is the digit looked at. i.e : 0.00X<- digit compared
            Assert.Equal(expected, actual,3);
        }
    }
}