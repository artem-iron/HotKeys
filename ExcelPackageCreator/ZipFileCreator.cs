using System;
using System.IO;
using System.IO.Compression;

namespace ExcelPackageCreator
{
    public class ZipFileCreator : Parameters
    {
        public static void HandleZipCreation(char keyChar)
        {
            if (keyChar == TRIGGER_KEY)
            {
                string zipName = $"{ZIP_NAME}{DateTime.Now:yy-MM-dd-HH-mm-ss}{ZIP_EXTENSION}";

                bool zipIsCreated = CreateZipFile(DIRECTORY_PATH, zipName);

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
                File.Move(zipName, zipName.Replace(ZIP_EXTENSION, XLSX_EXTENSION));
            }
        }
    }
}
