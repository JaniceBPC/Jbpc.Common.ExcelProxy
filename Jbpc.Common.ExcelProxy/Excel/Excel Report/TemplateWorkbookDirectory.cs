using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Jbpc.Common.Excel
{
    public static class TemplateWorkbookDirectory
    {
        private static string TemplateReportsDirectoryName = "Template Reports";
        private static string TemplateReportsDirectory(DirectoryInfo directoryInfo)
        {
            var initialDirectory = directoryInfo;

            while (directoryInfo.Exists)
            {
                var folderNames = directoryInfo.GetDirectories().Select(x=> x.Name).ToList();

                if (folderNames.Contains(TemplateReportsDirectoryName))
                {
                    return Path.Combine(directoryInfo.FullName, TemplateReportsDirectoryName);
                }

                if (directoryInfo.Parent == null)
                {
                    throw new ApplicationException($"TemplateReports directory \"{TemplateReportsDirectoryName}\" not found in path: \"{initialDirectory.FullName}\"");
                }
                directoryInfo = directoryInfo.Parent;
            }
            throw new ApplicationException($"TemplateReports directory \"{TemplateReportsDirectoryName}\" not found in path: \"{initialDirectory.FullName}\"");
        }
        public static string PathName()
        {
            var entryAssembly = Assembly.GetEntryAssembly();


            if (entryAssembly == null)
            {
                throw new ApplicationException($"Unable to find entry assembly and thus unable to navigate to the template excel reports directory: \"{TemplateReportsDirectoryName}\" ");
            }
            var fileInfo = new FileInfo(entryAssembly.Location);

            var originalAssemblyPath= typeof(TemplateWorkbookDirectory).Assembly.CodeBase;

            originalAssemblyPath = Uri.UnescapeDataString(new UriBuilder(originalAssemblyPath).Path);

            var directory = Path.GetDirectoryName(originalAssemblyPath) ?? originalAssemblyPath;

            Console.WriteLine($"Entry assembly original path={originalAssemblyPath}, directory={directory}");

            return TemplateReportsDirectory(new DirectoryInfo(directory));
        }
    }
}
