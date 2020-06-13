using System;
using System.IO;
using System.Reflection;

namespace ServiceStatus.Helpers
{
    class CommonHelper
    {
        public string GetProjectPath()
        {
            var assemblyPath = Assembly.GetCallingAssembly().CodeBase;
            var projectPath = new Uri(assemblyPath).LocalPath;
            return Path.GetDirectoryName(projectPath) + "\\";
        }
    }
}
