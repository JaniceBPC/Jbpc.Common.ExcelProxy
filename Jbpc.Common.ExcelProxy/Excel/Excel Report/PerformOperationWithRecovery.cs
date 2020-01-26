using System;

namespace Jbpc.Common.Excel
{
    public static class PerformOperationWithRecovery
    {
        private const int MAX_DEPTH = 4;

        public static void PerformOperation(Action lambda) => RecursivePerformOperationWithRecovery(lambda);

        private static void RecursivePerformOperationWithRecovery(Action lambda, int depth = 0)
        {
            var application = OfficeApplication.Singleton;

            if (depth < MAX_DEPTH)
            {
                try
                {
                    lambda();
                }
                catch (ApplicationException ex)
                {
                    throw ex;
                }
                catch (Exception)
                {
                    application.ReleaseApplication();

                    RecursivePerformOperationWithRecovery(lambda, depth+1);
                }
            }
        }
    }
}
