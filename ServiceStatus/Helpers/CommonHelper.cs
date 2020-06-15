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

        public enum YesNo
        {
            yes=1,
            Yes = 1,
            no =0,
            No = 0
        }
    }
}
