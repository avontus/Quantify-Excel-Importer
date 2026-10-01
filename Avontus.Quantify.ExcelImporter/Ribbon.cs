using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Avontus.Rental.Library;
using Avontus.Rental.Library.Security;
using Avontus.Quantify.ExcelImporter.Utils;
using System.Reflection;
using Avontus.Rental.Library.Settings;

namespace Avontus.Quantify.ExcelImporter
{
    public partial class Ribbon
    {
        #region Import Estimate 

        private static int _headerRow;
        private static int _partNumberCol;
        private static int _descrptionCol;
        private static int _quantityCol;
        private static int _weightCol;
        private static InformationList _informationList = InformationList.NewInformationList();

        private void AvontusRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // For use when merging the code in to import an estimate into Quantify
            //this.chkIgnorePNWithNoQty.Checked = Properties.Settings.Default.SkipBlankPNWithQty;
        }

        private void chkIgnorePNWithNoQty_Click(object sender, RibbonControlEventArgs e)
        {
            //Properties.Settings.Default.SkipBlankPNWithQty = chkIgnorePNWithNoQty.Checked;
            //Properties.Settings.Default.Save();
        }

        private void btnQuantifyExport_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (QuantifyInstalled() == false)
                {
                    MessageBoxHelper.Show("Quantify is not installed on this computer (cannot find the file 'Avontus.Rental.UI.exe')", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _informationList = InformationList.NewInformationList();
                ExportFile(true);
                LaunchQuantify();
            }
            catch (Exception ex)
            {
                MessageBoxHelper.Show("Error while opening BOM in Quantify: " + ex.Message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnSaveToQuantifyFile_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _informationList = InformationList.NewInformationList();
                ExportFile(false);
            }
            catch (Exception ex)
            {
                MessageBoxHelper.Show("Error while saving Quantify export file: " + ex.Message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void ExportFile(bool openInQuantify)
        {

            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null)
            {
                MessageBoxHelper.Show("This workbook is currently in protected mode. Please enable editing to continue.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Get my docs folder, exit if not found
            string tempPath = QuantifyMyDocsFolder();

            if (tempPath == "")
            {
                MessageBoxHelper.Show("An unexpected error occurred trying to reference your My Documents' folder. Please contact your system administrator to correct this problem", "Invalid Path", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // If page not valid then exit, method does the warning
            if (PageIsValid() == false)
            {
                MessageBoxHelper.Show("Spreadsheet does not appear to be valid. Expect columns:\r\n" +
                    "English: Product Code, Part Number, Description, Quantity, Qty, Weight\r\n" +
                    "Italian: Codice, Descrizione, Q.ta', Peso U.\r\n" +
                    "Some cells may be merged.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Set the column numbers
            SetColumns();

            if (openInQuantify)
            {

                // File name if attaching to an estimate
                string excelFile = "";

                if (Globals.QuantifyImporter.Application.ActiveWorkbook.Name != null &&
                    Globals.QuantifyImporter.Application.ActiveWorkbook.Name.Length > 0)
                {
                    if (Globals.QuantifyImporter.Application.ActiveWorkbook.Saved == false)
                    {
                        var res = MessageBoxHelper.Show("This file has not been saved and will not be attached to estimate. Would you like to save this file?", "Quantify Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (res == DialogResult.Yes)
                        {
                            Globals.QuantifyImporter.Application.ActiveWorkbook.Save();
                        }
                    }

                    // Check again if saved from above
                    if (Globals.QuantifyImporter.Application.ActiveWorkbook.Saved)
                    {
                        // Get the directory of the temp folder
                        excelFile = Path.Combine(Path.GetTempPath(), Globals.QuantifyImporter.Application.ActiveWorkbook.Name);

                        if (File.Exists(excelFile))
                        {
                            // Delete, ignore failure
                            try
                            {
                                File.Delete(excelFile);
                            }
                            catch
                            {

                            }
                        }

                        // Create copy of file
                        Globals.QuantifyImporter.Application.ActiveWorkbook.SaveCopyAs(excelFile);

                        WriteBomToFile(tempPath + "\\temp.txt", tempPath + "\\sd.txt", excelFile);
                    }
                    else
                    {
                        WriteBomToFile(tempPath + "\\temp.txt", tempPath + "\\sd.txt", "");
                    }

                    if (_informationList.Count > 0)
                    {
                        DialogResult result = DialogResult.Yes;
                        result = MessageBoxHelper.Show("The spreadsheet was converted but there are errors in the export.\r\n\r\nWould you like to view these errors?", "Quantify Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            using (frmInformationList frm = new frmInformationList(_informationList))
                            {
                                frm.ShowDialog();
                            }
                        }
                    }
                }
            }
            else
            {
                using (SaveFileDialog dlg = new SaveFileDialog())
                {

                    // Get file name from file create dialog box
                    dlg.Title = "Quantify Import File";
                    dlg.Filter = "Quantify Import File (*.qimport)|*.qimport";

                    if (dlg.ShowDialog() == DialogResult.Cancel)
                        return;

                    string filePath = dlg.FileName;

                    bool saved = WriteBomToFile(tempPath + "\\temp.txt", filePath, filePath);

                    if (_informationList.Count == 0)
                    {
                        MessageBoxHelper.Show("Successfully saved BOM to export file", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Utils.ReportUsage.FileSaved();
                    }
                    else
                    {
                        DialogResult res = DialogResult.Yes;
                        if (saved)
                            res = MessageBoxHelper.Show("Successfully saved materials to export file but there were errors.\r\n\r\nWould you like to view these errors?", "Quantify Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        else
                            res = MessageBoxHelper.Show("There were errors and the file could not be saved.\r\n\r\nWould you like to view these errors?", "Quantify Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (res == DialogResult.Yes)
                        {
                            using (frmInformationList frm = new frmInformationList(_informationList))
                            {
                                frm.ShowDialog();
                            }
                        }
                    }
                }
            }
        }

        private bool WriteBomToFile(string tempFile, string exportFile, string attachFile)
        {

            // Delete file if it exists
            if (File.Exists(tempFile))
                File.Delete(tempFile);

            // If it still exists then warn and exit
            if (File.Exists(tempFile))
            {
                _informationList.AddFailureItem(string.Format("Cannot overwrite file: {0}. Please delete manually and retry.", tempFile));
                return false;
            }

            bool lineWritten = false;

            using (StreamWriter sw = File.CreateText(tempFile))
            {

                // Attach qimport
                if (attachFile.Length > 0)
                {
                    // Write qimport file to attach
                    sw.WriteLine("VisioFile:" + attachFile);
                }
                else
                {
                    // Else write blank file lines
                    sw.WriteLine("VisioFile:");
                }

                Excel.Worksheet ws = (Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet;
                int i = _headerRow + 1;
                int blankRowCount = 0;
                while (i < 10000)
                {
                    Excel.Range range = ws.get_Range("A" + i.ToString(), "M" + i.ToString());
                    System.Array row = (System.Array)range.Cells.Value;

                    // If blank row then increment, else reset
                    if (FlatList(row) == "")
                        blankRowCount++;
                    else
                        blankRowCount = 0;

                    // Set part number
                    string partNumber = "";
                    if (((object[,])(row))[1, _partNumberCol] != null)
                        partNumber = ((object[,])(row))[1, _partNumberCol].ToString();

                    // Set qty
                    string qty = "";
                    if (((object[,])(row))[1, _quantityCol] != null)
                        qty = ((object[,])(row))[1, _quantityCol].ToString();

                    // See if qty is numeric
                    int tryQty;
                    if (qty.Length > 0 && int.TryParse(qty, out tryQty) == false && !partNumber.ToUpper().StartsWith("TOTAL "))
                        _informationList.AddWarningItem("Item " + partNumber + " on row " + i.ToString() + " has an invalid quantity of '" + qty + "' and was skipped");
                    else if (partNumber.Length == 0 && qty.Length > 0) // && chkIgnorePNWithNoQty.Checked == false)
                        _informationList.AddWarningItem("Row " + i.ToString() + " has a quantity but no product code and was skipped");

                    // See if part number is blank
                    if (partNumber.Length > 0 && !partNumber.ToUpper().StartsWith("TOTAL ") && qty.Length > 0)
                    {
                        string line = String.Format("{0}\t{1}\t0\t", ((object[,])(row))[1, _partNumberCol].ToString(), ((object[,])(row))[1, _quantityCol].ToString());
                        sw.WriteLine(line);
                        lineWritten = true;
                    }

                    i++;

                    if (blankRowCount > 15)
                        break;
                }
            }


            if (lineWritten == false)
            {
                _informationList.AddFailureItem("There are no items found to write to the file.");
                return false;
            }

            // Encrypt and delete original file
            bool result = FileBOM.Process(tempFile, exportFile, "");

            try
            {
                File.Delete(tempFile);
            }
            catch
            {
                // Ignore
            }

            return true;
        }

        private bool QuantifyInstalled()
        {
            string quantifyExe = "";

            // Get the path for Quantify
            string quantifyPath = "";

            try
            {
                // Get the registry key for where Quantify is installed
                RegEdit reg = new RegEdit(Microsoft.Win32.Registry.LocalMachine, @"Software\Avontus");

                if (reg != null)
                    quantifyPath = reg.Read("Quantify");

                // If null then try 6432 node
                if (reg == null || quantifyPath == null)
                {
                    reg = new RegEdit(Microsoft.Win32.Registry.LocalMachine, @"Software\Wow6432Node\Avontus");
                    // Try and grab the 64-bit registry
                    if (reg != null)
                        quantifyPath = reg.Read("Quantify");
                }

                if (quantifyPath == null)
                    return false;

                if (quantifyPath.EndsWith(@"\") == false)
                    quantifyPath += "\\";

                // Set path to Quantify exe
                quantifyExe = quantifyPath + "Avontus.Rental.UI.exe";

                // Warn if Quantify not found
                if (File.Exists(quantifyExe))
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

        private void LaunchQuantify()
        {
            string quantifyExe = "";

            // Get the path for Quantify
            string quantifyPath = "";

            try
            {
                // Get the registry key for where Quantify is installed
                RegEdit reg = new RegEdit(Microsoft.Win32.Registry.LocalMachine, @"Software\Avontus");

                if (reg != null)
                    quantifyPath = reg.Read("Quantify");

                // If null then try 6432 node
                if (reg == null || quantifyPath == null)
                {
                    reg = new RegEdit(Microsoft.Win32.Registry.LocalMachine, @"Software\Wow6432Node\Avontus");
                    // Try and grab the 64-bit registry
                    if (reg != null)
                        quantifyPath = reg.Read("Quantify");
                }

                if (quantifyPath.EndsWith(@"\") == false)
                    quantifyPath += "\\";

                // Set path to Quantify exe
                quantifyExe = quantifyPath + "Avontus.Rental.UI.exe";

                // Warn if Quantify not found
                if (!File.Exists(quantifyExe))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBoxHelper.Show("Unable to launch Quantify: " + ex.Message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            try
            {
                ProcessStartInfo pi = new ProcessStartInfo();
                pi.WorkingDirectory = quantifyPath;
                pi.FileName = quantifyExe;
                Process.Start(pi);
                Utils.ReportUsage.QuantifyLaunched();
            }
            catch (Exception ex)
            {
                MessageBoxHelper.Show("Error launching Quantify: " + ex.Message, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static void SetColumns()
        {
            Excel.Worksheet ws = (Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet;
            Excel.Range range = ws.get_Range("A" + _headerRow.ToString(), "M" + _headerRow.ToString());
            System.Array row = (System.Array)range.Cells.Value;

            int col = 1;
            foreach (string item in row)
            {
                if (item != null)
                {
                    switch (item.Replace(" ", "").ToUpper())
                    {
                        case "CODICE":
                        case "PARTNUMBER":
                        case "PRODUCTCODE":
                            _partNumberCol = col;
                            break;
                        case "DESCRIPTION":
                        case "DESCRIZIONE":
                            _descrptionCol = col;
                            break;
                        case "QUANTITY":
                        case "QTY":
                        case "Q.TA'":
                            _quantityCol = col;
                            break;
                        case "WEIGHT":
                        case "PESOU.":
                            _weightCol = col;
                            break;
                    }
                }
                col++;
            }
        }

        private static bool PageIsValid()
        {

            // If workbook or sheet is null then exit
            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null ||
                Globals.QuantifyImporter.Application.ActiveSheet == null)
                return false;
            else
            {
                Excel.Worksheet ws = (Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet;

                // Look in the first 20 rows
                int i = 1;
                while (i < 20)
                {
                    if (IsHeaderRow(ws, i))
                        return true;
                    i++;
                }

                return false;
            }
        }

        private static bool IsHeaderRow(Excel.Worksheet ws, int i)
        {
            Excel.Range range = ws.get_Range("A" + i.ToString(), "M" + i.ToString());
            System.Array row = (System.Array)range.Cells.Value;

            string rowString = FlatList(row);
            if (
                (rowString.Contains("CODICE") ||
                 rowString.Contains("PRODUCT CODE") ||
                 rowString.Contains("PART NUMBER") ||
                 rowString.Contains("PRODUCT CODE") ||
                 rowString.Contains("PART NUMBER"))
                &&
                (rowString.Contains("DESCRIPTION") ||
                 rowString.Contains("DESCRIZIONE"))
                &&
                (rowString.Contains("QUANTITY") ||
                 rowString.Contains("QTY") ||
                 rowString.Contains("Q.TA'"))
                &&
                (rowString.Contains("WEIGHT") ||
                 rowString.Contains("PESO U."))
               )
            {
                _headerRow = i;
                return true;
            }
            else
                return false;
        }

        private static string FlatList(System.Array arr)
        {
            string flattened = "";
            foreach (object item in arr)
            {
                if (item != null)
                    flattened += item.ToString();
            }

            return flattened.ToUpper();
        }
        /// <summary>
        /// This is a special method that looks for an environment variable that can be defined
        /// by IT admins on citrix and RDP that redirects to a path that they can specify in
        /// this folder instead of My Documents.
        /// 
        /// This method issues dialog boxes
        /// </summary>
        /// <returns>Returns the environment variable validated path, or My Documents.</returns>
        private static string QuantifyMyDocsFolder()
        {
            string tempPath = "";
            // First attempt to get Quantify_Path environment variable. We have
            // this documented for larger customers that use citrix
            if (Environment.GetEnvironmentVariable("Quantify_Path") != null)
            {
                tempPath = Environment.GetEnvironmentVariable("Quantify_Path");
                if (Directory.Exists(tempPath) == false)
                {
                    MessageBoxHelper.Show("The environment variable QUANTIFY_PATH set by your system administrator points to an invalid location. Please contact your system administrator to correct this problem", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                // Get path to my docs folder and append text file
                tempPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (Directory.Exists(tempPath) == false)
                {
                    MessageBoxHelper.Show("An unexpected error occurred trying to reference your My Documents' folder. Please contact your system administrator to correct this problem", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return "";
                }
            }

            return tempPath;
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Process.Start("https://docs.avontus.com/display/QUAN/Working+with+Avontus+Quantify+Importer"); 
                Utils.ReportUsage.AvontusAssist();
            }
            catch
            {
                //ignore
            }
        }

        private void btnCheckForUpdates_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string path = System.AppDomain.CurrentDomain.BaseDirectory;
                path = Path.Combine(path, "updater.exe");
                if (File.Exists(path))
                {
                    Process.Start(path, "/checknow", "", null, "");

                    // Report usage
                    Utils.ReportUsage.UpdateCheck();
                }
            }
            catch
            {
                // Ignore
            }
        }
        #endregion

        #region Generalized Importer

        private void btnImportFromExcel_Click(object sender, RibbonControlEventArgs e)
        {
            GeneralizedImporter importer = new GeneralizedImporter();
            
            importer.ImportSheet(); 
        }

        #endregion

        private void btnCreateTemplates_Click(object sender, RibbonControlEventArgs e)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            GeneralizedImporter importer = new GeneralizedImporter();
            importer.CreateTemplates(version); 
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            string server = "";
            var ctx = Rental.Library.Utility.DbUtil.GetConnectionManager();
            if (CommonConfigurationSettings.IsUsingDataPortal)
                server = CommonConfigurationSettings.RealDataPortalName.Replace("WcfPortal.svc", "");
            else
                server = CommonConfigurationSettings.SqlServerName; //ctx.Connection.DataSource;
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            MessageBox.Show("Avontus Quantify Excel Importer"
                + "\r\n\r\nServer: " + server
                + "\r\n\r\nVersion: " + version
                + "\r\n\r\nCopyright 2018, Avontus Software Corporation.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnWorkbookValidation_Click(object sender, RibbonControlEventArgs e)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            GeneralizedImporter importer = new GeneralizedImporter();
            importer.WorkbookValidation(version);
        }
    }
}
