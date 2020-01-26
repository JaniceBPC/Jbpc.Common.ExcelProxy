using System.IO;
using System.Reflection;

namespace Jbpc.Common.Excel
{
    public static class TemplateWorkbookDirectory
    {
        public static string PathName()
        {
            var entryAssembly = Assembly.GetEntryAssembly();
            var fileInfo = new FileInfo(entryAssembly.Location);
            return Path.GetFullPath(Path.Combine(fileInfo.DirectoryName, @"..\..\..\Template Reports"));
        }
    }
}
