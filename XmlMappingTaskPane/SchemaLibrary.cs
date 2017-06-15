//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Diagnostics;
using System.Globalization;
using System.Security;
using Microsoft.Win32;

namespace XmlMappingTaskPane
{
    static class SchemaLibrary
    {
        /// <summary>
        /// Get the alias for the schema.
        /// </summary>
        /// <param name="strNamespace">A string specifying the root namespace of the schema.</param>
        /// <param name="intLCID">An integer specifying the current UI language in LCID format.</param>
        /// <returns></returns>
        public static string GetAlias(string strNamespace, int intLCID)
        {
            try
            {
                //try to get the HKLM hive
                RegistryKey regHKLMKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", false);

                if (regHKLMKey != null && string.Equals(regHKLMKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKLMKey.OpenSubKey(@"Alias", false);
                    string strHKLMName = (string)regAlias.GetValue(intLCID.ToString(CultureInfo.InvariantCulture), string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKLMName))
                        return strHKLMName;

                    //check for a culture-invariant one
                    strHKLMName = (string)regAlias.GetValue("0", string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKLMName))
                        return strHKLMName;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to use HKLM: " + ex.Message);
            }

            try
            {
                //HKLM was no good, try HKCU
                RegistryKey regHKCUKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", false);

                if (regHKCUKey != null && string.Equals(regHKCUKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKCUKey.OpenSubKey(@"Alias", false);

                    if (regAlias == null)
                        return string.Empty;

                    string strHKCUName = (string)regAlias.GetValue(intLCID.ToString(CultureInfo.InvariantCulture), string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKCUName))
                        return strHKCUName;

                    //check for a culture-invariant one
                    strHKCUName = (string)regAlias.GetValue("0", string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKCUName))
                        return strHKCUName;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to use HKCU: " + ex.Message);
            }

            return string.Empty;
        }

        /// <summary>
        /// Set the alias for the schema
        /// </summary>
        /// <param name="strNamespace">A string specifying the root namespace of the schema.</param>
        /// <param name="strValue">A string specifying the alias.</param>
        /// <param name="intLCID">An integer specifying the current UI language in LCID format.</param>
        /// <returns>True if the alias was saved in the registry, False otherwise.</returns>
        public static bool SetAlias(string strNamespace, string strValue, int intLCID)
        {
            try
            {
                //HKLM was no good, try HKCU
                RegistryKey regHKCUKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", true);

                if (regHKCUKey == null)
                {
                    //create it
                    regHKCUKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    regHKCUKey.SetValue("Key", strNamespace, RegistryValueKind.String);
                }


                if (regHKCUKey != null && string.Equals(regHKCUKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKCUKey.OpenSubKey(@"Alias", true);

                    if (regAlias == null)
                    {
                        regAlias = regHKCUKey.CreateSubKey(@"Alias");
                    }

                    regAlias.SetValue(intLCID.ToString(CultureInfo.InvariantCulture), strValue, RegistryValueKind.String);
                    return true;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to write to HKCU: " + ex.Message);
            }

            return false;
        }
    }
}
