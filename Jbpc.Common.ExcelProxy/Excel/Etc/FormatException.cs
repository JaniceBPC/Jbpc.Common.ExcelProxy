using System;
using System.Diagnostics;
using System.Text;

namespace Jbpc.Common.Excel
{
    public static class FormatException
    {
        public static string Format(Exception ex, bool includeStackTrace = true, string message = "")
        {
            var sb = new StringBuilder();

            if (message == "") message = ex.Message != "" ? ex.Message : "No Message";

            if (message != "") sb.AppendLine($"{message}{Environment.NewLine}");

            sb.AppendLine(FormatNestedExceptions(ex, includeStackTrace));

            return sb.ToString();
        }
        private static string FormatNestedExceptions(Exception ex, bool includeStackTrace)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"{Environment.NewLine}{Divider}");
            sb.AppendLine($"Type of exception: {ex.GetType().Name}");
            sb.AppendLine($"Exception Message: {ex.Message}");

            if (includeStackTrace)
                sb.AppendLine(FormatStackTrace.PrettyPrintStackFrame(new StackTrace(ex, true)));

            if (ex.InnerException != null)
                sb.AppendLine(FormatNestedExceptions(ex.InnerException, includeStackTrace));
            
            return sb.ToString();
        }
        private static string Divider => "-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_";

    }
}
