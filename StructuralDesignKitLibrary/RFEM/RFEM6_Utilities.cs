using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dlubal.WS.Rfem6.Application;
using Dlubal.WS.Rfem6.Model;
using System.IO;
using System.ServiceModel;
using static System.Net.Mime.MediaTypeNames;
using ApplicationClient = Dlubal.WS.Rfem6.Application.RfemApplicationClient;
using ModelClient = Dlubal.WS.Rfem6.Model.RfemModelClient;
using Dlubal.RFEM5;
using System.Net;
using System.ServiceModel.Channels;

namespace StructuralDesignKitLibrary.RFEM
{
    public class RFEM6_Utilities : IRFEM_Utilities_Interface<ModelClient>
    {


        private BasicHttpBinding Binding
        {
            get
            {
                //Define the basic http binding 
                BasicHttpBinding binding = new BasicHttpBinding
                {
                    // Send timeout is set to 180 seconds.
                    SendTimeout = new TimeSpan(0, 0, 180),
                    UseDefaultWebProxy = true,
                    //Limiting the return value to 1 GByte
                    MaxReceivedMessageSize = 1000000000,

                };

                return binding;
            }
        }

        public ApplicationClient GetRFEMApplication()
        {
            ApplicationClient application = null;
            string CurrentDirectory = Directory.GetCurrentDirectory();

            try
            {
                application_information ApplicationInfo;
                try
                {

                    //define where RFEM is listening
                    EndpointAddress Address = new EndpointAddress("http://localhost:8081");

                    // connects to RFEM6 or RSTAB9 application
                    application = new ApplicationClient(Binding, Address);

                }
                catch (Exception exception)
                {
                    if (application != null)
                    {
                        if (application.State != CommunicationState.Faulted)
                        {
                            application.Close();
                        }
                        else
                        {
                            application.Abort();
                        }

                        application = null;
                    }
                }
                return application;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void CloseRFEMModel(ModelClient model)
        {
            throw new NotImplementedException();
        }

        public ModelClient GetActiveModel()
        {

            var RFEMApp = GetRFEMApplication();
            var activeModelString = RFEMApp.get_active_model();


            ModelClient model = new ModelClient(Binding, new EndpointAddress(activeModelString));

            return model;
        }

        public List<ModeShape> GetAllStandardizedDisplacement(ModelClient model, int NVC)
        {
            throw new NotImplementedException();
        }

        public List<FeMeshNode> GetFENodes(ModelClient model, int NVC)
        {
            throw new NotImplementedException();
        }

        public List<double> GetModalMasses(ModelClient model, int NVC)
        {
            throw new NotImplementedException();
        }

        public List<double> GetNaturalFrequencies(ModelClient model, int NVC)
        {
            throw new NotImplementedException();
        }


        public ModelClient OpenModel()
        {
            throw new NotImplementedException();
        }
    }
}
