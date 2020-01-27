using Microsoft.Office.Interop.Excel;

namespace Jbpc.Common.Excel
{
    public class OfficeApplication
    {
        private static OfficeApplication singleton = new OfficeApplication();
        public static OfficeApplication Singleton => singleton;
        private Application application;
        public Application Application
        {
            get
            {
                if (application != null) return application;

                PerformOperationWithRecovery.PerformOperation(InstantiateApplication);

                return application;
            }
        }
        private void InstantiateApplication()
        {
            application = new ApplicationClass();
        }
        public void ReleaseApplication()
        {
            if (application != null)
            {
                try
                {
                    application.Quit();
                }
                finally
                {
                    application = null;
                    KillExcelProcess.KillHeadlessExcelProcesses();
                }
            }
        }
    }
}
