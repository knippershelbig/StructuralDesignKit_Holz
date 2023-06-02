using Dlubal.DynamPro;
using Dlubal.RFEM5;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System;

namespace StructuralDesignKitLibrary.RFEM
{
	public interface IRFEM_Utilities_Interface<T> where T : class
	{

		T GetActiveModel();

		T OpenModel();



		void CloseRFEMModel(T model);



		/// <summary>
		/// Get the modal masses from RFEM as list (one per mode)
		/// </summary>
		/// <param name="model">RFEM Model</param>
		/// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
		/// <returns>Returns a list of double representing the modal masses</returns>
		List<double> GetModalMasses(T model, int NVC);


		/// <summary>
		/// Get the natural frequencies calculated in RF-DynamPro
		/// </summary>
		/// <param name="model">RFEM Model</param>
		/// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
		/// <returns>Returns a list of double representing the natural frequencies calculated in RF-DynamPro</returns>
		List<double> GetNaturalFrequencies(T model, int NVC);



		/// <summary>
		/// Get all the standardized displacement for the different mode shapes and FE nodes
		/// </summary>
		/// <param name="model">RFEM Model</param>
		/// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
		/// <returns></returns>
		List<ModeShape> GetAllStandardizedDisplacement(T model, int NVC);


		/// <summary>
		/// Return a list of all FE nodes in the model
		/// </summary>
		/// <param name="model">RFEM Model</param>
		/// <param name="NVC">NaturalVibrationCase index defined in RF-DynamPro</param>
		/// <returns></returns>
		List<FeMeshNode> GetFENodes(T model, int NVC);



	}
}