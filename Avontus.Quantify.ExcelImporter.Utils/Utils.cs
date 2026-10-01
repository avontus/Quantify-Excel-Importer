using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography;
using SmartAssembly.ReportUsage;

namespace Avontus.Quantify.ExcelImporter.Utils
{
    public static class ReportUsage
    {
        public static void FileSaved()
        {
            SmartAssembly.ReportUsage.UsageCounter.ReportUsage("File Saved");
        }
        public static void QuantifyLaunched()
        {
            SmartAssembly.ReportUsage.UsageCounter.ReportUsage("Quantify Launched");
        }
        public static void UpdateCheck()
        {
            SmartAssembly.ReportUsage.UsageCounter.ReportUsage("Check for Update");
        }
        public static void AvontusAssist()
        {
            SmartAssembly.ReportUsage.UsageCounter.ReportUsage("Avontus Assist");
        }
    }

    public static class FileBOM
    {
        private readonly static string _item = @"321yeKym";

        public static bool Process(string inputFile, string outputFile, string keyFile)
        {
            try
            {
                if (keyFile != "")
                    throw new FileNotFoundException();

                UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = UE.GetBytes(_item);

                string cryptFile = outputFile;
                using (FileStream fsCrypt = new FileStream(cryptFile, FileMode.Create))
                {
                    RijndaelManaged RMCrypto = new RijndaelManaged();
                    using (CryptoStream cs = new CryptoStream(fsCrypt, RMCrypto.CreateEncryptor(key, key), CryptoStreamMode.Write))
                    {
                        using (FileStream fsIn = new FileStream(inputFile, FileMode.Open))
                        {
                            int data;
                            while ((data = fsIn.ReadByte()) != -1)
                                cs.WriteByte((byte)data);
                        }
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
            
        }

    }
}
