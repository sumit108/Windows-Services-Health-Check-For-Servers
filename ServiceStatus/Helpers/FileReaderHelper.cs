using System;
using System.IO;
using System.Reflection;

namespace ServiceStatus
{
    public class FileHelper
    {
        public string GetProjectPath()
        {
            var assemblyPath = Assembly.GetCallingAssembly().CodeBase;
            var projectPath = new Uri(assemblyPath).LocalPath;
            return Path.GetDirectoryName(projectPath) + "\\";
        }
    }
}

