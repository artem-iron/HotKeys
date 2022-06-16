using System;
using System.IO;
using System.IO.Compression;
using System.Configuration;

namespace ExcelPackageCreator
{
    public class ZipFileCreator
    {

        private static readonly Configuration _config = GetConfig();
        private static readonly char triggerKey = '';

        public static string MonitoredPath
        {
            get
            {
                return _config.AppSettings.Settings["MonitoredPath"].Value;
            }
        }
        public static string SavePath
        {
            get
            {
                return _config.AppSettings.Settings["SavePath"].Value;
            }
        }

        private static string ZipName
        {
            get
            {
                return _config.AppSettings.Settings["ZipName"].Value;
            }
        }
        private static string ZipExtension
        {
            get
            {
                return _config.AppSettings.Settings["ZipExtension"].Value;
            }
        }
        private static string XlsxExtension
        {
            get
            {
                return _config.AppSettings.Settings["XlsxExtension"].Value;
            }
        }

        public static void HandleZipCreation(char keyChar)
        {
            if (keyChar == triggerKey)
            {
                string zipName = $"{SavePath}{ZipName}{DateTime.Now:yy-MM-dd-HH-mm-ss}{ZipExtension}";

                bool zipIsCreated = CreateZipFile(MonitoredPath, zipName);

                ChangeExtension(zipIsCreated, zipName);
            }
        }

        private static bool CreateZipFile(string directoryPath, string resultingFileName)
        {
            if (!Directory.Exists(directoryPath))
            {
                return false;
            }

            ZipFile.CreateFromDirectory(directoryPath, resultingFileName);

            return true;
        }

        private static void ChangeExtension(bool zipIsCreated, string zipName)
        {
            if (zipIsCreated && File.Exists(zipName))
            {
                File.Move(zipName, zipName.Replace(ZipExtension, XlsxExtension));
            }
        }

        public static void StoreNewMonitoredPath(string newPath)
        {
            _config.AppSettings.Settings["MonitoredPath"].Value = newPath;

            _config.Save(ConfigurationSaveMode.Modified);
        }

        public static void StoreNewSavePath(string newPath)
        {
            _config.AppSettings.Settings["SavePath"].Value = newPath + (newPath.EndsWith("\\") ? "" : "\\");

            _config.Save(ConfigurationSaveMode.Modified);
        }
        private static Configuration GetConfig()
        {
            string applicationName = Environment.GetCommandLineArgs()[0];

            string exePath = System.IO.Path.Combine(Environment.CurrentDirectory, applicationName);

            return ConfigurationManager.OpenExeConfiguration(exePath);
        }
    }
}
