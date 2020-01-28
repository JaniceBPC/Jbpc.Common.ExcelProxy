using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Jbpc.Common.Excel
{
    public static class TemplateWorkbookDirectory
    {
        private static string TemplateReportsDirectoryName = "Template Reports";
        private static string TemplateReportsDirectory(string path)
        {
            var initialPath = path;

            while (new DirectoryInfo(path).Exists)
            {
                var folderNames = new DirectoryInfo(path).GetDirectories().Select(x=> x.Name).ToList();

                if (folderNames.Contains(TemplateReportsDirectoryName))
                {
                    return Path.Combine(path, TemplateReportsDirectoryName);
                }

                path = new DirectoryInfo(path).Parent?.FullName ?? "";
            }
            throw new ApplicationException($"TemplateReports directory \"{TemplateReportsDirectoryName}\" not found in path: \"{initialPath}\"");
        }
        public static string PathName()
        {
            var entryAssembly = Assembly.GetEntryAssembly();

            if (entryAssembly == null)
            {
                throw new ApplicationException($"Unable to find entry assembly and thus unable to navigate to the template excel reports directory: \"{TemplateReportsDirectoryName}\" ");
            }
            var fileInfo = new FileInfo(entryAssembly.Location);

            var codeBase = typeof(TemplateWorkbookDirectory).Assembly.CodeBase;

            var uri = new UriBuilder(codeBase);
            var path = Uri.UnescapeDataString(uri.Path);

            return TemplateReportsDirectory(fileInfo.DirectoryName);
        }
    }
}
