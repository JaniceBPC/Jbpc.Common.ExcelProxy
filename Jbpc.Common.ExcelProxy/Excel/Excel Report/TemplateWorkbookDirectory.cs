using System;
using System.IO;
using System.Reflection;

namespace Jbpc.Common.Excel
{
    public static class TemplateWorkbookDirectory
    {
        public static string PathName()
        {
            var codeBase = typeof(TemplateWorkbookDirectory).Assembly.CodeBase;

            var uri = new UriBuilder(codeBase);
            var path = Uri.UnescapeDataString(uri.Path);

            var entryAssembly = Assembly.GetEntryAssembly();
            var fileInfo = new FileInfo(entryAssembly.Location);

            return Path.GetFullPath(Path.Combine(fileInfo.DirectoryName, @"..\..\..\TemplateReports"));
        }
    }
}
