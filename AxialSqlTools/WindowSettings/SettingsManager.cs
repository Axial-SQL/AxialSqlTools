using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace AxialSqlTools
{
    class SettingsManager
    {

        //private static byte[] Protect(byte[] data)
        //{
        //    try
        //    {
        //        // Use the current user scope to encrypt the data.
        //        return ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
        //    }
        //    catch (CryptographicException e)
        //    {
        //        Console.WriteLine($"A cryptographic error occurred: {e.Message}");
        //        return null;
        //    }
        //}

        //private static byte[] Unprotect(byte[] data)
        //{
        //    try
        //    {
        //        // Decrypt the data using the current user scope.
        //        return ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
        //    }
        //    catch (CryptographicException e)
        //    {
        //        Console.WriteLine($"A cryptographic error occurred: {e.Message}");
        //        return null;
        //    }
        //}


        private static RegistryKey GetRoot()
        {
            var settingKeyRoot = Registry.CurrentUser.CreateSubKey(@"AxialSqlTools");
            var settingsKey = settingKeyRoot.CreateSubKey("Settings");

            return settingsKey;
        }

        private static string GetRegisterValue(string parameter)
        {
            try
            {
                using (var rootKey = GetRoot())
                {
                    var value = rootKey.GetValue(parameter);
                    return value?.ToString() ?? string.Empty;
                }
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private static bool SaveRegisterValue(string parameterName, string parameterValue)
        {
            try
            {
                using (var rootKey = GetRoot())
                {
                    rootKey.SetValue(parameterName, parameterValue);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // ----------------------------------------------------------------------
        public static string GetMyEmail()
        {
            return GetRegisterValue("MyEmail");
        }
        public static bool SaveMyEmail(string myEmail)
        {
            return SaveRegisterValue("MyEmail", myEmail);
        }


        public static string GetTemplatesFolder()
        {
            var folder = GetRegisterValue("ScriptTemplatesFolder");

            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AxialSqlToolsTemplates");

                SaveTemplatesFolder(folder);

                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }

            return folder;
        }

        public static bool SaveTemplatesFolder(string folder)
        {
            return SaveRegisterValue("ScriptTemplatesFolder", folder);
        }

    }
}
