using System;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avontus.Quantify.ExcelImporter
{
    /// <summary>
    /// A useful class to read/write/delete/count registry keys
    /// </summary>
    public class RegEdit : IDisposable
    {
        public void Dispose()
        {
            if (baseRegistryKey != null)
            {
                baseRegistryKey.Dispose();
                baseRegistryKey = null;
            }
        }

        public RegEdit()
        {
#if DEBUG
            showError = true;
#endif
        }

        public RegEdit(RegistryKey baseRegKey, string subKey)
            : this()
        {
            this.BaseRegistryKey = baseRegKey;
            this.SubKey = subKey;
        }

        private bool showError = false;
        /// <summary>
        /// A property to show or hide error messages 
        /// (default = false)
        /// </summary>
        public bool ShowError
        {
            get
            {
                return showError;
            }
            set
            {
                showError = value;
            }
        }

        private string subKey = String.Format("Software\\{0}\\{1}", Application.CompanyName, Application.ProductName);
        /// <summary>
        /// A property to set the SubKey value
        /// </summary>
        public string SubKey
        {
            get
            {
                return subKey;
            }
            set
            {
                subKey = value;
            }
        }

        private RegistryKey baseRegistryKey = Registry.CurrentUser;
        /// <summary>
        /// A property to set the BaseRegistryKey value.
        /// (default = Registry.LocalMachine)
        /// </summary>
        public RegistryKey BaseRegistryKey
        {
            get
            {
                return baseRegistryKey;
            }
            set
            {
                baseRegistryKey = value;
            }
        }


        /// <summary>
        /// To read a registry key.
        /// input: KeyName (string)
        /// output: value (string) 
        /// </summary>
        public string Read(string keyName)
        {
            // Opening the registry key
            RegistryKey rk = baseRegistryKey;
            // Open a subKey as read-only
            RegistryKey sk1 = rk.OpenSubKey(subKey);
            if (keyName.Contains("\\"))
            {
                var tmp = keyName.Right(keyName.Length - keyName.LastIndexOf('\\') - 1);
                sk1 = sk1.OpenSubKey(keyName.Replace(tmp, ""));
                keyName = tmp;
            }
            // If the RegistrySubKey doesn't exist -> (null)
            if (sk1 == null)
            {
                return null;
            }
            else
            {
                try
                {
                    if (!sk1.GetValueNames().Contains(keyName) && sk1.GetSubKeyNames().Contains(keyName))
                    {
                        sk1 = sk1.OpenSubKey(keyName);
                        return (string)sk1.GetValue(sk1.GetValueNames().First());

                    }
                    else if (sk1.ValueCount != 0)
                        return (string)sk1.GetValue(keyName);
                    return null;
                }
                catch (Exception e)
                {
                    // AAAAAAAAAAARGH, an error!
                    ShowErrorMessage("Error reading registry key ", keyName, e);

                    return null;
                }
            }
        }




        /// <summary>
        /// To write into a registry key.
        /// input: KeyName (string) , Value (object)
        /// output: true or false 
        /// </summary>
        public bool Write(string keyName, object value)
        {
            try
            {
                // Setting
                RegistryKey rk = baseRegistryKey;
                // I have to use CreateSubKey 
                // (create or open it if already exits), 
                // 'cause OpenSubKey open a subKey as read-only
                RegistryKey sk1 = rk.CreateSubKey(subKey);
                // Save the value
                sk1.SetValue(keyName, value);

                return true;
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error writing Key ", keyName, e);
                return false;
            }
        }


        /// <summary>
        /// To delete a registry key.
        /// input: KeyName (string)
        /// output: true or false 
        /// </summary>
        public bool DeleteKey(string keyName)
        {
            try
            {
                // Setting
                RegistryKey rk = baseRegistryKey;
                RegistryKey sk1 = rk.CreateSubKey(subKey);
                // If the RegistrySubKey doesn't exists -> (true)
                if (sk1 == null || !sk1.GetSubKeyNames().Contains(keyName))
                    return true;
                else
                    sk1.DeleteValue(keyName);

                return true;
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error deleting SubKey ", subKey, e);
                return false;
            }
        }


        /// <summary>
        /// To delete a sub key and any child.
        /// input: void
        /// output: true or false 
        /// </summary>
        public bool DeleteSubKeyTree()
        {
            try
            {
                // Setting
                RegistryKey rk = baseRegistryKey;
                RegistryKey sk1 = rk.OpenSubKey(subKey);
                // If the RegistryKey exists, I delete it
                if (sk1 != null)
                    rk.DeleteSubKeyTree(subKey);

                return true;
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error deleting SubKey ", subKey, e);
                return false;
            }
        }


        /// <summary>
        /// Retrieve the count of subkeys at the current key.
        /// input: void
        /// output: number of subkeys
        /// </summary>
        public int SubKeyCount()
        {
            try
            {
                // Setting
                RegistryKey rk = baseRegistryKey;
                RegistryKey sk1 = rk.OpenSubKey(subKey);
                // If the RegistryKey exists...
                if (sk1 != null)
                    return sk1.SubKeyCount;
                else
                    return 0;
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error retrieving keys of ", subKey, e);
                return 0;
            }
        }


        public IEnumerable<RegistryKey> GetSubKeys(bool readOnly = true)
        {
            IEnumerable<RegistryKey> subKeys = new List<RegistryKey>();
            try
            {
                RegistryKey rk = baseRegistryKey;
                RegistryKey sk1 = readOnly ? rk.OpenSubKey(subKey) : rk.CreateSubKey(subKey);
                // If the RegistryKey exists...
                if (sk1 != null)
                {
                    if (readOnly)
                        subKeys = sk1.GetSubKeyNames().Select(x => sk1.OpenSubKey(x));
                    else
                        subKeys = sk1.GetSubKeyNames().Select(x => sk1.CreateSubKey(x));
                }
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error retrieving keys of ", subKey, e);
            }
            return subKeys;
        }


        /// <summary>
        /// Retrieve the count of values in the key.
        /// input: void
        /// output: number of keys
        /// </summary>
        public int ValueCount()
        {
            try
            {
                // Setting
                RegistryKey rk = baseRegistryKey;
                RegistryKey sk1 = rk.OpenSubKey(subKey);
                // If the RegistryKey exists...
                if (sk1 != null)
                    return sk1.ValueCount;
                else
                    return 0;
            }
            catch (Exception e)
            {
                ShowErrorMessage("Error retrieving keys of ", subKey, e);
                return 0;
            }
        }

        private void ShowErrorMessage(string message, string subKey, Exception e)
        {
            MessageBoxHelper.Show(message + " " + subKey + ": " + e.Message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
