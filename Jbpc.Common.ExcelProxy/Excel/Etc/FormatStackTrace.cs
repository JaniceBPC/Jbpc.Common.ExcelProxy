using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace Jbpc.Common.Excel
{
    public static class FormatStackTrace
    {
        public static string PrettyPrintStackFrame(StackTrace stackTrace)
        {
            var sb = new StringBuilder();

            var frames = new List<StackFrame>(stackTrace.GetFrames());

            foreach (var stackFrame in frames)
            {
                var methodType = stackFrame.GetMethod().ReflectedType;

                if (methodType != null)
                    sb.AppendLine($"{methodType.FullName} {stackFrame.GetFileName()} {stackFrame.GetFileLineNumber()}");
            }

            var msg = sb.ToString();
            return sb.ToString();
        }
    }
}
