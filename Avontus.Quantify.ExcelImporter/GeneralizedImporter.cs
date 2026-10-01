using Avontus.Core;
using Avontus.Rental.Library;
using Avontus.Rental.Library.Request;
using Avontus.Rental.Library.Security;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Avontus.Quantify.ExcelImporter
{
    class GeneralizedImporter
    {
        /// <summary>
        /// GlobalKeyType is used by the Global value import page
        /// </summary>
        private enum GlobalKeyType
        {
            Project = 1,
            Priority = 2,
            Shift = 3,
            Driver = 4,
            CountyParish = 5,
            State = 6,
            Step = 7,
            Manufacturer = 8,
            RequestList1 = 9,
            RequestList2 = 10,
            RequestList3 = 11,
            JobList1 = 12,
            JobList2 = 13,
            ShipmentList1 = 14,
            ShipmentList2 = 15,
            Invalid = 99
        }

        private enum BoolResult
        {
            False = 0,
            True = 1,
            Empty = 2,
            Invalid = 10
        }

        private enum CycleType  // like ArrearsBillingCycleType with an added None value
        {
            Daily = 0,
            Monthly = 1,
            None = 10
        }

        private enum InvoiceOption
        {                           // Quantify Equivalent value
            SingleOrder = 0,        // InvoiceOrderGroupType.PerJob
            IndividualOrder = 1,    // InvoiceOrderGroupType.PerShipmentOrScaffoldTag
            GroupByPO = 2,          // InvoiceOrderGroupType.PerPO
            None = 10               // No equivalent in Quantify
        }

        private class OrderInfo
        {
            public Order Order;
            public bool IsWorkOrder;
            public bool IsExisting;

            public OrderInfo(Order order, bool isWorkOrder, bool isExisting)
            {
                Order = order;
                IsWorkOrder = isWorkOrder;
                IsExisting = isExisting;
            }
        }

        internal class ImportBalance
        {
            public Guid StockingLocationID;
            public Guid ParentLocationID;
            public Guid ParentTradingPartnerID;
            public string RentStartDate;
            public int WorksheetColumn;
            public Guid ScaffoldTagActivityTypeID;
            public bool IsBillable;
            public Guid OrderID;
            public List<Tuple<Guid, double>> Balances;

            public ImportBalance(Guid stockingLocationiD, Guid parentLocationID, Guid parentTradingPartnerID,
                string rentStartDate, int worksheetColumn)
            {
                StockingLocationID = stockingLocationiD;
                ParentLocationID = parentLocationID;
                ParentTradingPartnerID = parentTradingPartnerID;
                RentStartDate = rentStartDate;
                WorksheetColumn = worksheetColumn;
                Balances = new List<Tuple<Guid, double>>();
            }

            public void Add(Guid productID, double quantity)
            {
                this.Balances.Add(new Tuple<Guid, double>(productID, quantity));
            }
        }

        internal class BranchImportBalance
        {
            public Guid StockingLocationID;
            public int WorksheetColumn;
            public List<Tuple<Guid, double>> Balances;

            public BranchImportBalance(Guid stockingLocationiD, int worksheetColumn)
            {
                StockingLocationID = stockingLocationiD;
                WorksheetColumn = worksheetColumn;
                Balances = new List<Tuple<Guid, double>>();
            }

            public void Add(Guid productID, double quantity)
            {
                this.Balances.Add(new Tuple<Guid, double>(productID, quantity));
            }
        }
        internal delegate string NormSting(string s);
        NormSting NormStr = ImporterData.NormStr;

        // Colors used - specified as RGB decimal values or color names
        System.Drawing.Color warningColor = System.Drawing.Color.FromArgb(255, 248, 220);
        System.Drawing.Color lightBlue = System.Drawing.Color.FromArgb(222, 236, 247);
        System.Drawing.Color lightGray = System.Drawing.Color.FromArgb(217, 217, 217);
        System.Drawing.Color lighterGray = System.Drawing.Color.FromArgb(242, 242, 242);
        System.Drawing.Color requiredInterior = System.Drawing.Color.FromArgb(0, 99, 0); // a dark green
        System.Drawing.Color requiredFontColor = System.Drawing.Color.White;

        const string NonImportSheetID = "Not an import sheet";

        const string Validated = "Validated.";
        const string Imported = "Imported.";
        const string Results = "Results";
        const int HeaderRow = 6;

        string[] ValidGlobalListKeys = { "PROJECT", "PRIORITY", "STEP", "SHIFTS", "DELIVERYDRIVERS", "COUNTIES/PARISHES",
                "STATES/PROVINCES", "MANUFACTURERS", "REQUESTLIST1","REQUESTLIST2","REQUESTLIST3",
                "JOBLIST1", "JOBLIST2", "SHIPMENTLIST1", "SHIPMENTLIST2"};

        string[] ValidScaffoldJobListNames = { "SCAFFOLDTAGACTIVITYTYPE", "SCAFFOLDACTIVITYLIST1", "SCAFFOLDACTIVITYLIST2", "SCAFFOLDACTIVITYLIST3",
            "SCAFFOLDLIST1", "SCAFFOLDLIST2", "SCAFFOLDLIST3", "SCAFFOLDLIST4",
            "SCAFFOLDLIST5", "SCAFFOLDLIST6", "SCAFFOLDLIST7"};

        string[] ValidScaffoldTagStatuses = { "ESTIMATE", "ON RENT", "OFF RENT", "COMPLETED",
            "CUSTOM 1", "CUSTOM 2", "CUSTOM 3", "CUSTOM 4", "CUSTOM 5", "CUSTOM 6" };

        string[] ValidBillingMethods = { "ARREARS", "FATA", "ADVANCE" };

        string[] ValidInvoiceOptions = { "SINGLE ORDER", "INDIVIDUAL ORDER", "GROUP BY PO" };


        private enum ProgressDisplayType
        {
            Row = 1,
            Column = 2,
            Shipment = 3,
            Adjustment = 4
        }

        private const int maxBrokenRulesReported = 3;
        private Tuple<int, int> notFound = new Tuple<int, int>(0, 0);

        string versionStr = string.Format("[version: {0}]", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

        #region Importer Methods

        internal void ImportSheet()
        {
            const string resultHeader = Results;

            // If workbook or sheet is null then exit
            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null ||
                Globals.QuantifyImporter.Application.ActiveSheet == null)
            {
                return;
            }

            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null)
            {
                MessageBoxHelper.Show("This workbook is currently in protected mode. Please enable editing to continue.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Login
            var principal = Avontus.Core.ApplicationContext.User;
            AvontusUser user = null;
            if (principal == null || !principal.Identity.IsAuthenticated)
            {
                frmLogin loginForm = new frmLogin();
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    loginForm.Username = ""; // set locally when debugging
                    loginForm.Password = ""; // set locally when debugging
                }

                DialogResult dlgResult = loginForm.ShowDialog();
                if (dlgResult == DialogResult.Cancel)
                    return;

                AvontusPrincipal.Logout();
                try
                {
                    AvontusPrincipal.Login(loginForm.Username, loginForm.Password);
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Cannot open database"))
                    {
                        MessageBoxHelper.Show("Unable to connect to the database.  ", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        MessageBoxHelper.Show(string.Format("Database connection error.  Exweption: (0). ", ex.Message),
                            "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                // If login failed
                if (!AvontusPrincipal.CurrentIdentity.IsAuthenticated)
                {
                    MessageBoxHelper.Show("User Name or Password is invalid", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // If not admin
                user = AvontusUser.GetUser(loginForm.Username);
                if (user.IsAdministrator == false)
                {
                    MessageBoxHelper.Show("You must be an administrator to use the Quantify Excel Importer.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            // Exit if connected to dataportal
            if (Rental.Library.Settings.CommonConfigurationSettings.IsUsingDataPortal)
            {
                MessageBoxHelper.Show("The import cannot proceed when connected to a remote server. Connect to a local server before importing.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Get the dimension of the worksheet
            var ws = (Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet;

            var activeRange = (object[,])ws.UsedRange.Value;
            if (activeRange == null)
            {
                MessageBoxHelper.Show("The worksheet appears to be empty;  nothing to import", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }

            int maxRows = activeRange.GetLength(0);

            if (maxRows < 3)
            {
                MessageBoxHelper.Show("The workhseet must have at least three rows -- Import Identifier, Column Headers and at least one data row", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }

            int maxColumns = activeRange.GetLength(1);
            ImportConfig.ImportSheetType sheetType = ImportConfig.ImportSheetType.Unknown;

            string importIdentifier = FindImportIdentifer(ws, maxRows, maxColumns, out sheetType);
            if (importIdentifier == "")
            {
                MessageBoxHelper.Show("The workhseet must have a cell in the header containing the Import Identifier ('Avontus Import:  <Valid Import Type>'", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }

            if (sheetType == ImportConfig.ImportSheetType.Normalized)
            {
                ImportConfig importConfig = ImportConfig.GetImportConfig(importIdentifier);
                if (importConfig == null)
                {
                    MessageBoxHelper.Show(string.Format("The Import Identifier in this worksheet {0} is not recognized. Can't import this data.", importIdentifier), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return;
                }

                string headerIssues = VerifyRequiredHeaders(ws, importConfig, maxRows, maxColumns, importIdentifier, resultHeader);
                if (headerIssues != "")
                {
                    MessageBoxHelper.Show(headerIssues, "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int resultCol = SetResultColumn(ws, maxColumns, HeaderRow, resultHeader);

                RemoveColumnBackgroundColors(ws, importConfig, HeaderRow, maxRows, resultCol);

                bool brokenRulesFound = PerformImport(ws, importConfig, maxRows, maxColumns, HeaderRow, resultCol, user);
            }
            else if (importIdentifier == "IMPORT:NONSTANDARD1")
            {
                bool result = NonStandardImport1(ws, maxRows, maxColumns);

            }
            else if (sheetType == ImportConfig.ImportSheetType.BranchOrLaydownBalance)
            {
                bool allFound = false;

                Tuple<int, int> nameLabel = notFound;
                Tuple<int, int> branchOrLaydown = notFound;
                Tuple<int, int> partNumber = notFound;
                Tuple<int, int> description = notFound;

                int resultRow = 0;

                for (int row = maxRows; row >= 11; row--)
                {
                    string data = CellData(ws, row, 1);
                    if (data == Results)
                    {
                        resultRow = row;
                        maxRows--;
                        break;
                    }
                    else if (data != "")
                    {
                        resultRow = row + 1;
                        WriteCell(resultRow, 1, ws, Results);
                        break;
                    }

                }
                for (int col = 2; col <= maxColumns; col++)
                {
                    if (CellData(ws, resultRow, col) != Imported)
                    {
                        WriteCell(resultRow, col, ws, "");
                    }
                }

                for (int row = 1; row <= maxRows; row++)
                {
                    for (int col = 1; col <= maxColumns; col++)
                    {
                        var cellRange = ((Excel.Range)ws.Cells[row, col]);

                        if (cellRange.Formula.ToString() != "")
                        {
                            string value = NormStr(cellRange.Formula.ToString());
                            if (!(importIdentifier == "IMPORT:BRANCHORLAYDOWNBALANCES"))
                            {
                                return;  // THIS IS AN EXIT POINT
                            }

                            switch (value)
                            {
                                case "NAME:":
                                    nameLabel = new Tuple<int, int>(row, col);
                                    break;
                                case "BRANCHORLAYDOWN:":
                                    branchOrLaydown = new Tuple<int, int>(row, col);
                                    break;
                                case "PARTNUMBER":
                                    partNumber = new Tuple<int, int>(row, col);
                                    break;
                                case "DESCRIPTION":
                                    description = new Tuple<int, int>(row, col);
                                    break;
                            }

                            if (nameLabel != notFound &&
                                branchOrLaydown != notFound &&
                                partNumber != notFound &&
                                description != notFound &&
                                resultRow != 0)
                            {
                                allFound = true;
                                break;
                            }
                        }
                    }
                    if (allFound)
                        break;
                }

                if (resultRow == 0)
                {
                    resultRow = maxRows + 2;
                    WriteCell(resultRow, 1, ws, Results, useColor: true, color: warningColor);
                }
                else
                {
                    WriteCell(resultRow + 2, 1, ws, "", useColor: true); // Clear text and remove color
                    for (int col = nameLabel.Item2 + 1; col < maxColumns; col++)
                    {
                        WriteCell(resultRow, col, ws, useColor: true); // Clear interior color
                    }
                }

                // Check rules for Balance Import sheet 
                bool brokenRulesReported = false;

                if (partNumber == notFound)
                {
                    string msg = string.Format("Missing heading - \"{0}\" ", "Part Number");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }
                if (description == notFound)
                {
                    string msg = string.Format("Missing heading - \"{0}\" ", "Description");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (nameLabel == notFound)
                {
                    string msg = string.Format("Missing label - \"{0}\" ", "Name:");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (branchOrLaydown == notFound)
                {
                    string msg = string.Format("Missing label - \"{0}\" ", "Branch or Laydown: ");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (nameLabel == notFound ||
                    branchOrLaydown == notFound ||
                    partNumber == notFound ||
                    description == notFound ||
                    resultRow == 0)
                {
                    string msg = string.Format("Validation terminated due to worksheet formatting problems identified. ");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                }

                if (!brokenRulesReported)
                {
                    ImportBranchBalances(ws, maxRows, maxColumns, resultRow);
                }
            }
            else if (sheetType == ImportConfig.ImportSheetType.ScaffoldTagBalance ||
                     sheetType == ImportConfig.ImportSheetType.JobSiteBalance)
            {
                bool allFound = false;

                Tuple<int, int> rentStart = notFound;
                Tuple<int, int> tagLabel = notFound;
                Tuple<int, int> jobLabel = notFound;
                Tuple<int, int> partNumber = notFound;
                Tuple<int, int> description = notFound;
                Tuple<int, int> activityLabel = notFound;
                Tuple<int, int> summaryLabel = notFound;

                bool tagBalanceImport = false;
                int resultRow = 0;


                for (int row = maxRows; row > 9; row--)
                {
                    string data = CellData(ws, row, 1);
                    if (data == Results)
                    {
                        resultRow = row;
                        WriteCell(resultRow, 1, ws, Results);
                        break;
                    }
                    else if (!EmptyRow(ws, row, maxColumns + 1))
                    {
                        resultRow = row + 1;
                        WriteCell(resultRow, 1, ws, Results);
                        break;
                    }
                }
                // Clear old messages
                for (int col = 2; col <= maxColumns; col++)
                {
                    if (CellData(ws, resultRow, col) != Imported)
                    {
                        WriteCell(resultRow, col, ws, "");
                    }
                }

                for (int row = 1; row <= maxRows; row++)
                {
                    for (int col = 1; col <= maxColumns; col++)
                    {
                        var cellRange = ((Excel.Range)ws.Cells[row, col]);
                        if (cellRange.Formula.ToString() != "")
                        {
                            string value = NormStr(cellRange.Formula.ToString());
                            if (importIdentifier == "IMPORT:SCAFFOLDTAGBALANCES")
                            {
                                tagBalanceImport = true;
                            }
                            else if (!(importIdentifier == "IMPORT:JOBSITEBALANCES"))
                            {
                                return;  // THIS IS AN EXIT POINT
                            }

                            switch (value)
                            {
                                case "RENTSTART:":
                                    rentStart = new Tuple<int, int>(row, col);
                                    break;
                                case "TAG:":
                                    tagLabel = new Tuple<int, int>(row, col);
                                    break;
                                case "JOBSITE:":
                                    jobLabel = new Tuple<int, int>(row, col);
                                    break;
                                case "PARTNUMBER":
                                    partNumber = new Tuple<int, int>(row, col);
                                    break;
                                case "DESCRIPTION":
                                    description = new Tuple<int, int>(row, col);
                                    break;
                                case "ACTIVITY:":
                                    activityLabel = new Tuple<int, int>(row, col);
                                    break;
                                case "SUMMARY:":
                                    summaryLabel = new Tuple<int, int>(row, col);
                                    break;
                            }

                            if (rentStart != notFound &&
                                (tagLabel != notFound || tagBalanceImport == false) &&
                                jobLabel != notFound &&
                                partNumber != notFound &&
                                description != notFound &&
                                resultRow != 0)
                            {
                                allFound = true;
                                break;
                            }
                        }
                    }
                    if (allFound)
                        break;
                }

                if (resultRow == 0)
                {
                    resultRow = maxRows + 2;
                    WriteCell(resultRow, 1, ws, Results, useColor: true, color: warningColor);
                }
                else
                {
                    WriteCell(resultRow + 2, 1, ws, "", useColor: true); // Clear text and remove color
                    for (int col = rentStart.Item2 + 1; col < maxColumns; col++)
                    {
                        WriteCell(resultRow, col, ws, useColor: true); // Clear interior color
                    }
                }

                // Check rules for Balance Import sheet 
                bool brokenRulesReported = false;


                if (rentStart == notFound)
                {
                    string msg = string.Format("Missing label - \"{0}\" ", "Rent Start:");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (partNumber == notFound)
                {
                    string msg = string.Format("Missing heading - \"{0}\" ", "Part Number");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }
                if (description == notFound)
                {
                    string msg = string.Format("Missing heading - \"{0}\" ", "Description");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (tagBalanceImport && tagLabel == notFound)
                {
                    string msg = string.Format("Missing label - \"{0}\" ", "Tag:");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (!tagBalanceImport && jobLabel == notFound)
                {
                    string msg = string.Format("Missing label - \"{0}\" ", "Job Site: ");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                    brokenRulesReported = true;
                }

                if (rentStart == notFound ||
                    (tagLabel == notFound && tagBalanceImport == true) ||
                    jobLabel == notFound ||
                    partNumber == notFound ||
                    description == notFound ||
                    resultRow == 0)
                {
                    string msg = string.Format("Validation terminated due to worksheet formatting problems identified. ");
                    WriteCell(resultRow + 2, 1, ws, msg, useColor: true, append: true, color: warningColor);
                }

                if (!brokenRulesReported)
                {
                    string label = tagBalanceImport ? "Tag " : "Job ";
                    int tagOrJobRow = tagBalanceImport ? tagLabel.Item1 : jobLabel.Item1;

                    if (rentStart.Item1 != (jobLabel.Item1 - 1))
                    {
                        WriteCell(resultRow + 2, 1, ws, "Job Site row must immediately follow Rent Start Row. ",
                            useColor: true, append: true, color: warningColor);
                        brokenRulesReported = true;
                    }

                    if (tagBalanceImport && (jobLabel.Item1 != (tagLabel.Item1 - 1)))
                    {
                        WriteCell(resultRow + 2, 1, ws, "Tag row must immediately follow Job Stie Row. ",
                            useColor: true, append: true, color: warningColor);
                        brokenRulesReported = true;
                    }

                    if (partNumber.Item1 != description.Item1 || partNumber.Item2 != (description.Item2 - 1))
                    {
                        WriteCell(resultRow + 2, 1, ws, "Description must follow Part Number on the same row. ", useColor: true, append: true, color: warningColor);
                        brokenRulesReported = true;
                    }

                    if (partNumber.Item1 <= tagOrJobRow)
                    {
                        WriteCell(resultRow + 2, 1, ws, string.Format("Part Number row must follow {0} row. ", label),
                            useColor: true, append: true, color: warningColor);
                        brokenRulesReported = true;
                    }

                    if (brokenRulesReported)
                    {
                        WriteCell(resultRow + 2, 1, ws, "Import aborted due to broken rules. ", useColor: true, append: true, color: warningColor);
                    }
                }

                if (!brokenRulesReported)
                {
                    ImportBalances(ws, maxRows, maxColumns, resultRow, tagBalanceImport,
                        rentStart, tagLabel, jobLabel, partNumber, description,
                        activityLabel, summaryLabel);
                }
            }
        }

        private string VerifyRequiredHeaders(Excel.Worksheet ws, ImportConfig importConfig,
            int maxRows, int maxColumns, string importIdentifier, string resultHeader)
        {
            bool hiddenColumnsReported = false;
            StringBuilder result = new StringBuilder();
            int maxConfiguredColumn = 0;

            foreach (var importCol in importConfig.ImportColumns)
            {
                if (importCol.Column > maxConfiguredColumn)
                {
                    maxConfiguredColumn = importCol.Column;
                }

                if (importCol.Column > maxColumns)
                {
                    result.Append(string.Format("Missing Column: {0}; ", importCol.ColumnHeading));
                    break;
                }
                else
                {
                    string expectedHeader = importCol.ColumnHeading;
                    string contents = CellData(ws, 6, importCol.Column);
                    if (NormStr(contents) != NormStr(expectedHeader))
                    {
                        result.Append(string.Format("Column(s) in this worksheet have been modified. Column {0} is missing or out of order. Correct the worksheet and retry. ", expectedHeader));
                        break;
                    }

                    if ((bool)((Excel.Range)ws.Cells[6, importCol.Column]).Columns.Hidden == true && !hiddenColumnsReported)
                    {
                        for (int row = 7; row <= maxRows; row++)
                        {
                            if (CellData(ws, row, importCol.Column) != "")
                            {
                                string msg = "One or more columns in the worksheet are hidden and contain data. Unhide all columns. ";
                                result.Append(msg);
                                hiddenColumnsReported = true;

                                break;
                            }
                        }
                    }
                }
            }

            if (maxColumns > maxConfiguredColumn && result.ToString() == "")
            {
                for (var col = maxConfiguredColumn + 1; col <= maxColumns; col++)
                {

                    if (CellData(ws, 6, col) != "" && CellData(ws, 6, col) != resultHeader)
                    {
                        result.Append(string.Format("Column(s) in this worksheet have been modified. Column {0} needs to be removed. Correct the worksheet and retry. ",
                            CellData(ws, 6, col)));
                        break;
                    }
                }
            }



            string returnValue = result.ToString();
            return returnValue == "" ? "" : "WorkSheet column header verification failed: " + result.ToString() + " Correct error(s) and retry. ";
        }


        /// <summary>
        ///  Stores the worksheet columns in the config file for this import for later reference
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="config"></param>
        /// <param name="maxRows"></param>
        /// <param name="maxColumns"></param>
        /// <param name="headerRow"></param>
        /// <returns></returns>
        private bool SetDataColumns(Excel.Worksheet ws, ImportConfig config, int maxRows, int maxColumns,
             int headerRow)
        {
            bool columnsSet = true;
            List<string> headers = new List<string>();
            for (var col = 1; col <= maxColumns; col++)
            {
                var cellRange = ((Excel.Range)ws.Cells[headerRow, col]);
                if (cellRange.Formula.ToString() != "")
                {
                    string str = NormStr(cellRange.Value.ToString());
                    headers.Add(str);
                }
            }

            for (var j = 0; j < config.ImportColumns.Count; j++)
            {
                string configHeader = NormStr(config.ImportColumns[j].ColumnHeading);
                if (headers.Contains(configHeader))
                {
                    int index = headers.IndexOf(configHeader);
                    config.ImportColumns[j].Column = index;
                }
                else
                {
                    config.ImportColumns[j].Column = -1;
                    if (config.ImportColumns[j].IsRequired)
                    {
                        columnsSet = false;
                    }
                }
            }
            return columnsSet;
        }
        /// <summary>
        /// Finds the HeaderRow defined as the row containing all required column headers for 
        /// this import. It can contain other Headers, including headers not used for the import
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="config"></param>
        /// <param name="maxRows"></param>
        /// <param name="maxColumns"></param>
        /// <param name="importIdentifier"></param>
        /// <returns></returns>
        private int FindHeaderRow(Excel.Worksheet ws, ImportConfig config, int maxRows, int maxColumns, string importIdentifier)
        {
            int headerRow = -1;
            List<string> requiredHeaders = new List<string>();

            foreach (ImportConfig.ImportColumn c in config.ImportColumns.Where(x => x.IsRequired == true))
            {
                string head = NormStr(c.ColumnHeading);
                requiredHeaders.Add(head);
            }

            for (int row = 1; row <= maxRows; row++)
            {
                List<string> foundHeaders = new List<string>();
                for (int col = 1; col <= maxColumns; col++)
                {
                    var cellRange = ((Excel.Range)ws.Cells[row, col]);
                    if (cellRange.Formula.ToString() != "")
                    {
                        string str = NormStr(cellRange.Value.ToString());
                        foundHeaders.Add(str);
                    }
                }
                bool foundRequiredHeaders = true;
                foreach (string h in requiredHeaders)
                {
                    if (!foundHeaders.Contains(h))
                    {
                        foundRequiredHeaders = false;
                        break;
                    }
                }
                if (foundRequiredHeaders)
                {
                    headerRow = row;
                    break;
                }
            }
            return headerRow;
        }

        /// <summary>
        /// Search for a cell containing the string "Avontus Import: [importIdentifier]"
        /// where [import identifier] is a valid Import Identifier. 
        /// Case is not important and spaces are ignored.  Anyting in the cell after the column
        /// is the identifier.  The identifier shold not contain a colon. 
        /// The worksheet should not contain more than one identifier and at least one is required. 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="maxRows"></param>
        /// <param name="maxColumns"></param>
        /// <returns></returns>
        private string FindImportIdentifer(Excel.Worksheet ws, int maxRows, int maxColumns, out ImportConfig.ImportSheetType sheetType)
        {
            string result = "";
            var cellRange = ((Excel.Range)ws.Cells[3, 1]);
            if (cellRange.Formula.ToString() != "")
            {
                string str = cellRange.Value.ToString();
                if (NormStr(str).Left(7) == "IMPORT:")
                {
                    result = NormStr(str.Trim());
                }
            }

            // Identifiy non-normalized imports 
            switch (result)
            {
                case "IMPORT:BRANCHORLAYDOWNBALANCES":
                    sheetType = ImportConfig.ImportSheetType.BranchOrLaydownBalance;
                    break;
                case "IMPORT:SCAFFOLDTAGBALANCES":
                    sheetType = ImportConfig.ImportSheetType.ScaffoldTagBalance;
                    break;
                case "IMPORT:JOBSITEBALANCES":
                    sheetType = ImportConfig.ImportSheetType.JobSiteBalance;
                    break;
                case "IMPORT:NONSTANDARD1":
                    sheetType = ImportConfig.ImportSheetType.Unknown;
                    break;
                default:
                    sheetType = ImportConfig.ImportSheetType.Normalized;
                    break;
            }
            return result;
        }

        private int SetResultColumn(Excel.Worksheet ws, int maxColumns, int headerRow, string resultHeader)
        {
            int resultCol = -1;

            for (var i = maxColumns; i > 0; i--)
            {

                var contents = ((Excel.Range)ws.Cells[headerRow, i]).Value;
                if (contents != null && contents.ToString() == resultHeader)
                {
                    resultCol = i;
                    //((Excel.Range)ws.Cells[headerRow, resultCol]).Value = resultHeader;
                    WriteCell(headerRow, resultCol, ws, resultHeader, true, bold: true, color: warningColor);
                    break;
                }
                else if (contents != null && contents.ToString().Trim() != "")
                {
                    resultCol = i + 1;
                    //((Excel.Range)ws.Cells[headerRow, resultCol]).Value = resultHeader;
                    WriteCell(headerRow, resultCol, ws, resultHeader, true, bold: true, color: warningColor);
                    break;
                }
            }

            // ((Excel.Range)ws.Cells[headerRow, resultCol]).Interior.Color = warningColor;
            try
            {
                ((Excel.Range)ws.Cells[headerRow, resultCol]).Columns.AutoFit();
                ((Excel.Range)ws.Cells[headerRow, resultCol]).Columns.WrapText = true;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;  // Consume the exception but go on.
            }

            return resultCol;
        }

        /// <summary>
        /// import the rows in the spreadsheet
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="config"></param>
        /// <param name="maxRows"></param>
        /// <param name="maxColumns"></param>
        /// <param name="headerRow"></param>
        /// <param name="infoList"></param>
        /// <returns></returns>
        private bool PerformImport(Excel.Worksheet ws, ImportConfig config, int maxRows, int maxColumns,
            int headerRow, int resultCol, AvontusUser user)
        {
            bool brokenRulesFound = false;

            List<Tuple<object, object, object, int, int, object>> importedObjects = new List<Tuple<object, object, object, int, int, object>>();
            ScaffoldTagStatusList statusList = ScaffoldTagStatusList.GetScaffoldTagStatusList(ActiveStatus.Both);

            // The importerTuple contains a group of dictionaries with lookup information to support 
            object importerTuple = ImporterData.GetData(config.ImportIdentifier);
            HashSet<string> jobNamesInSheet = new HashSet<string>();
            HashSet<string> foundNames = new HashSet<string>();
            Dictionary<string, Guid> savedOrders = new Dictionary<string, Guid>();
            Dictionary<string, Guid> savedWorkOrders = new Dictionary<string, Guid>();
            Dictionary<string, OrderInfo> orderDic = new Dictionary<string, OrderInfo>();
            Dictionary<string, OrderInfo> workOrderDic = new Dictionary<string, OrderInfo>();
            Dictionary<string, RentalPeriodAdjustment> adjustmentDic = new Dictionary<string, RentalPeriodAdjustment>();
            Dictionary<string, List<int>> adjustmentGoodRule = new Dictionary<string, List<int>>();
            Dictionary<string, bool> tagsOnSheet = new Dictionary<string, bool>();
            Dictionary<Guid, UnitHourRateProfile> profileDict = new Dictionary<Guid, UnitHourRateProfile>();

            List<string> namesSeen = new List<string>();
            List<string> seenProducts = new List<string>();
            List<string> seenConsumables = new List<string>();

            // Pre iteration checks for Request import
            if (NormStr(config.ImportIdentifier) == "IMPORT:REQUEST")
            {
                if (!GlobalClientConfig.OptionsSettings.EnableRequestPortal)
                {
                    MessageBoxHelper.Show("The Global Option to track Requests is not checked; the import can't be done.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;        // THIS IS AN END POINT
                }
            }


            for (int row = headerRow + 1; row <= maxRows; row++)
            {
                ShowProgress(ws, ProgressDisplayType.Row, row, maxRows, versionStr, resultCol);

                // Skip rows that have already been imported
                var rowImportStatus = ((Excel.Range)ws.Cells[row, resultCol]).Value;
                if (rowImportStatus != null)
                {
                    string previousResult = rowImportStatus.ToString().Trim();
                    if (previousResult == "Imported.")
                    {
                        continue;
                    }
                }

                // Populate the ImprortDictionary that "knows" import column information for each worksheet column
                bool importFailed = false;
                bool nonEmptyRow = false;
                Dictionary<string, ImportConfig.ImportColumn> importDict = new Dictionary<string, ImportConfig.ImportColumn>();
                Dictionary<string, List<string>> userDic = new Dictionary<string, List<string>>();

                foreach (var importColumn in config.ImportColumns)
                {
                    if (importColumn.Column != -1)
                    {
                        var cellRange = ((Excel.Range)ws.Cells[row, importColumn.Column]);
                        if (cellRange.Formula.ToString() != "")
                        {
                            nonEmptyRow = true;
                        }
                        importDict.Add(NormStr(importColumn.ColumnHeading), importColumn);
                    }
                }
                if (nonEmptyRow)
                {
                    switch (NormStr(config.ImportIdentifier))
                    {
                        case "IMPORT:GLOBALLISTS":
                            importFailed = ImportGlobalLists(importDict, ws, config, row, resultCol, importedObjects);
                            break;

                        case "IMPORT:PRODUCTCATEGORY":
                            importFailed = ImportProductCategory(importDict, ws, config, row, resultCol, importedObjects);
                            break;

                        case "IMPORT:PRODUCTCATALOG":
                            importFailed = ImportProductCatalog(importDict, ws, config, row, resultCol, importedObjects, seenProducts);
                            break;

                        case "IMPORT:CONSUMABLESCATEGORY":
                            importFailed = ImportConsumablesCategory(importDict, ws, config, row, resultCol, importedObjects);
                            break;

                        case "IMPORT:CONSUMABLESCATALOG":
                            importFailed = ImportConsumablesCatalog(importDict, ws, config, row, resultCol, importedObjects, seenConsumables);
                            break;

                        case "IMPORT:RENTADJUSTMENT":
                            importFailed = ImportRentAdjustment(importDict, ws, config, row, resultCol,
                                adjustmentDic, adjustmentGoodRule, importedObjects);
                            break;

                        case "IMPORT:RATEPROFILE":
                            importFailed = ImportRateProfile(importDict, ws, config, row, resultCol, importedObjects);
                            break;

                        case "IMPORT:CUSTOMER":
                            importFailed = ImportCustomer(importDict, ws, config, row, resultCol, importedObjects, namesSeen);
                            break;

                        case "IMPORT:CUSTOMERCONTACT":
                            importFailed = ImportCustomerContact(importDict, ws, config, row, resultCol,
                                 importedObjects);
                            break;

                        case "IMPORT:ORDER":
                            importFailed = ImportOrder(importDict, ws, config, row, resultCol,
                                 orderDic, workOrderDic, importedObjects);
                            break;
                        case "IMPORT:VENDOR":
                            importFailed = ImportVendor(importDict, ws, config, row, resultCol,
                                 importedObjects);
                            break;

                        case "IMPORT:VENDORCONTACT":
                            importFailed = ImportVendorContact(importDict, ws, config, row, resultCol,
                                 importedObjects);
                            break;

                        case "IMPORT:TAXRATE":
                            importFailed = ImportTaxRate(importDict, ws, config, row, resultCol,
                                 importedObjects, foundNames);
                            break;

                        case "IMPORT:JOBSITE":
                            importFailed = ImportJobsite(importDict, ws, config, row, resultCol,
                                 importedObjects, jobNamesInSheet);
                            break;

                        case "IMPORT:REQUEST":
                            importFailed = ImportRequest(importDict, ws, config, row, resultCol, user, userDic,
                                 importedObjects);
                            break;
                        case "IMPORT:SCAFFOLDJOBLIST":
                            importFailed = ImportScaffoldJobList(importDict, ws, config, row, resultCol,
                                 importedObjects);
                            break;

                        case "IMPORT:SCAFFOLDTAG":
                            Tuple<Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, int>> tagImporterData =
                                (Tuple<Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, int>>)importerTuple;

                            importFailed = ImportScaffoldTag(importDict, ws, config, row, resultCol,
                                importedObjects, tagImporterData, savedOrders, savedWorkOrders, tagsOnSheet);
                            break;

                        case "IMPORT:SCAFFOLDTAGACTIVITY":
                            if (GlobalClientConfig.ActivationSettings.Edition != Avontus.Rental.Library.Licensing.Activation.SoftwareEditions.Industrial)
                            {
                                MessageBoxHelper.Show(string.Format("Scaffold Tag ActivitY import requires the Industrial edition of Quantify. Unable to import worksheet.", config.ImportIdentifier), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return true;
                            }
                            else
                            {
                                Tuple<Dictionary<string, Guid>, Dictionary<Tuple<Guid, string>, Guid>, Dictionary<Tuple<Guid, Int16, string>, Guid>> activityImporterData =
                                    (Tuple<Dictionary<string, Guid>, Dictionary<Tuple<Guid, string>, Guid>, Dictionary<Tuple<Guid, Int16, string>, Guid>>)importerTuple;

                                importFailed = ImportScaffoldTagActivity(importDict, ws, config, row, resultCol,
                                     importedObjects, activityImporterData);
                            }
                            break;

                        case "IMPORT:UNITOFMEASURE":
                            importFailed = ImportUnitOfMeasure(importDict, ws, config, row, resultCol,
                                importedObjects, foundNames);
                            break;

                        case "IMPORT:UNITHOURRATEPROFILE":
                            importFailed = ImportUnitHourRateProfile(importDict, ws, config, row, resultCol,
                                importedObjects, profileDict);
                            break;

                        case "IMPORT:MULTIPLIER":
                            importFailed = ImportMultiplier(importDict, ws, config, row, resultCol,
                                importedObjects);
                            break;

                        default:
                            MessageBoxHelper.Show(string.Format("The Import Identifier, {0}, is not recognized. Unable to import worksheet.", config.ImportIdentifier), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                    }

                    if (importFailed)
                        brokenRulesFound = true;
                }
            }
            ClearProgress(ws);

            // Save Validated ImportObjects in Quantify

            bool noImports = importedObjects.Count == 0;
            string message = string.Format("Validation completed. {0} {1}",
                brokenRulesFound ? "Validation errors were found as noted." : "No validation errors noted.",
                noImports ? "There are no items to import." : "Do you want to import the validated items?");

            DialogResult response = MessageBoxHelper.Show(message,
                "Excel Importer", noImports ? MessageBoxButtons.OK : MessageBoxButtons.YesNo,
                noImports ? MessageBoxIcon.Information : MessageBoxIcon.Question);
            if (!noImports && response == DialogResult.Yes)
            {
                SaveImportedData(config.ImportIdentifier, ws, importedObjects);
                MessageBoxHelper.Show("Import complete.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return brokenRulesFound;
        }

        /// <summary>
        /// Import a row from the spreadsheet
        /// </summary>
        /// <param name="importDict"></param>
        /// <param name="ws"></param>
        /// <param name="config"></param>
        /// <param name="row"></param>
        /// <param name="resultCol"></param>
        /// <param name="infoList"></param>
        /// <param name="tagList"></param>
        /// <returns></returns>
        private bool ImportGlobalLists(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
           ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            int listColumn = config.ImportColumns.Where(x => NormStr(x.ColumnHeading) == "LISTNAME").First().Column;

            string keyValue = "";
            string valueStr = "";
            bool isActive = false;
            string description = "";
            bool isStateOrCounty = false;

            Tuple<string, ImportConfig.ImportColunnType, int> keyData;
            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "LISTNAME":
                            keyData = CellData("LISTNAME", ws, row, importDict);
                            keyValue = keyData.Item1;
                            if (keyValue != null)
                            {
                                keyValue = NormStr(keyValue);
                                if (!ValidGlobalListKeys.Contains(keyValue))
                                {
                                    WriteCell(row, resultCol, ws, "List Name is not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            else
                            {
                                isStateOrCounty = keyValue == "STATE" || keyValue == "COUNTY/PARISH";
                            }
                            break;
                        case "VALUE/NAME":
                            keyData = CellData("VALUE/NAME", ws, row, importDict);
                            valueStr = keyData.Item1;
                            if (valueStr == null || valueStr == "")
                            {
                                WriteCell(row, resultCol, ws, "Value/Name is not provided. Unable to import. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "ISACTIVE":
                            keyData = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(keyData, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE or FALSE. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                isActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "DESCRIPTION":
                            keyData = CellData("DESCRIPTION", ws, row, importDict);
                            description = keyData.Item1 ?? "";
                            break;
                        default:
                            importFailure = true;
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject1 = null;

                GlobalKeyType keyType = GlobalKeyType.Invalid;

                switch (keyValue)
                {
                    case "PROJECT":
                        keyType = GlobalKeyType.Project;
                        var project = Project.NewProject();
                        project.Name = valueStr;
                        project.IsActive = isActive;
                        project.Description = description;

                        if (project.IsSavable)
                        {
                            importObject1 = project;
                        }
                        else
                        {
                            foreach (var brokenRule in project.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "PRIORITY":
                        keyType = GlobalKeyType.Priority;
                        var priority = Priority.NewPriority();
                        priority.Name = valueStr;
                        priority.IsActive = isActive;
                        priority.Description = description;

                        if (priority.IsSavable)
                        {
                            importObject1 = priority;
                        }
                        else
                        {
                            foreach (var brokenRule in priority.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "STEP":
                        keyType = GlobalKeyType.Step;
                        var step = ScaffoldTagStep.NewScaffoldTagStep();
                        step.Name = valueStr;
                        step.IsActive = isActive;
                        step.Description = description;

                        if (step.IsSavable)
                        {
                            importObject1 = step;
                        }
                        else
                        {
                            foreach (var brokenRule in step.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "SHIFTS":
                        keyType = GlobalKeyType.Shift;
                        var shift = Shift.NewShift();
                        shift.Name = valueStr;
                        shift.IsActive = isActive;
                        shift.Description = description;
                        if (shift.IsSavable)
                        {
                            importObject1 = shift;
                        }
                        else
                        {
                            foreach (var brokenRule in shift.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "DELIVERYDRIVERS":
                        keyType = GlobalKeyType.Driver;
                        var driver = Driver.NewDriver();
                        driver.Name = valueStr;
                        driver.IsActive = isActive;
                        driver.Description = description;
                        if (driver.IsSavable)
                        {
                            importObject1 = driver;
                        }
                        else
                        {
                            foreach (var brokenRule in driver.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }

                        break;
                    case "COUNTIES/PARISHES":
                        keyType = GlobalKeyType.CountyParish;
                        var county = County.NewCounty();
                        county.Name = valueStr;

                        if (county.IsSavable)
                        {
                            importObject1 = county;
                        }
                        else
                        {
                            foreach (var brokenRule in county.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "STATES/PROVINCES":
                        keyType = GlobalKeyType.State;
                        var state = State.NewState();
                        state.Name = valueStr;

                        if (state.IsSavable)
                        {
                            importObject1 = state;
                        }
                        else
                        {
                            foreach (var brokenRule in state.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "MANUFACTURERS":
                        keyType = GlobalKeyType.Manufacturer;
                        var mfg = Manufacturer.NewManufacturer();
                        mfg.Name = valueStr;
                        mfg.Description = description;

                        if (mfg.IsSavable)
                        {
                            importObject1 = mfg;
                        }
                        else
                        {
                            foreach (var brokenRule in mfg.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "REQUESTLIST1":
                        var listItem1 = GlobalRequestListItem.NewGlobalRequestListItem(GlobalRequestListNumberTypes.GlobalList1);
                        keyType = GlobalKeyType.RequestList1;
                        listItem1.Name = valueStr;
                        listItem1.IsActive = isActive;
                        listItem1.Description = description;

                        if (listItem1.IsSavable)
                        {
                            importObject1 = listItem1;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem1.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Request List1 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "REQUESTLIST2":
                        var listItem2 = GlobalRequestListItem.NewGlobalRequestListItem(GlobalRequestListNumberTypes.GlobalList2);
                        keyType = GlobalKeyType.RequestList2;
                        listItem2.Name = valueStr;
                        listItem2.IsActive = isActive;
                        listItem2.Description = description;

                        if (listItem2.IsSavable)
                        {
                            importObject1 = listItem2;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem2.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Request List2 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "REQUESTLIST3":
                        var listItem3 = GlobalRequestListItem.NewGlobalRequestListItem(GlobalRequestListNumberTypes.GlobalList3);
                        keyType = GlobalKeyType.RequestList3;
                        listItem3.Name = valueStr;
                        listItem3.IsActive = isActive;
                        listItem3.Description = description;

                        if (listItem3.IsSavable)
                        {
                            importObject1 = listItem3;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem3.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Request List3 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "JOBLIST1":
                        keyType = GlobalKeyType.JobList1;
                        var listItem4 = JobList1.NewJobList1();
                        listItem4.Name = valueStr;
                        listItem4.IsActive = isActive;
                        listItem4.Description = description;

                        if (listItem4.IsSavable)
                        {
                            importObject1 = listItem4;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem4.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Job List1 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;
                    case "JOBLIST2":
                        keyType = GlobalKeyType.JobList2;
                        var listItem5 = JobList2.NewJobList2();
                        listItem5.Name = valueStr;
                        listItem5.IsActive = isActive;
                        listItem5.Description = description;

                        if (listItem5.IsSavable)
                        {
                            importObject1 = listItem5;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem5.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Job List2 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;

                    case "SHIPMENTLIST1":
                        keyType = GlobalKeyType.ShipmentList1;
                        var listItem6 = GlobalList1.NewGlobalList1();
                        listItem6.Name = valueStr;
                        listItem6.IsActive = isActive;
                        listItem6.Description = description;

                        if (listItem6.IsSavable)
                        {
                            importObject1 = listItem6;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem6.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Shipment List1 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;

                    case "SHIPMENTLIST2":
                        keyType = GlobalKeyType.ShipmentList2;
                        var listItem7 = GlobalList2.NewGlobalList2();
                        listItem7.Name = valueStr;
                        listItem7.IsActive = isActive;
                        listItem7.Description = description;

                        if (listItem7.IsSavable)
                        {
                            importObject1 = listItem7;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem7.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Shipment List2 Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                        break;

                    default:
                        break;
                }

                if (!importFailure && keyType != GlobalKeyType.Invalid)
                {
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject1, keyType, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }

            }
            return importFailure;
        }

        private bool ImportProductCategory(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            ProductCategory category = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            bool numberProvided = false;
            Guid? parentCategoryID = null;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "PARENTNAME":
                            data = CellData("PARENTNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                // First see if the parent is imported previously in the sheet (not yet in the database)
                                try
                                {
                                    // Search importedObjects for the parent  Product Category that matches and is validated previously
                                    var id = importedObjects.Where(x => ((ProductCategory)x.Item1).Name == data.Item1).Select(x => ((ProductCategory)x.Item1).ProductCategoryID).Single();
                                    parentCategoryID = id;
                                }
                                catch
                                {
                                    // See if the databae has the Parent Product Category
                                    ProductCategory parent = ProductCategory.GetProductCategory(data.Item1, ProductType.Product);
                                    if (category == null || category.Name.ToUpper() != data.Item1.ToUpper())
                                    {

                                        WriteCell(row, resultCol, ws, "Parent category doesn't exist in the database.  Cant validate this row. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        parentCategoryID = category.ProductCategoryID;
                                    }
                                }
                            }
                            break;
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                category = ProductCategory.GetProductCategory(data.Item1, ProductType.Product);
                                if (category != null && category.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Product category exists; can't import this category. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    category = ProductCategory.NewProductCategory(ProductType.Product);
                                    category.Name = data.Item1;
                                    if (parentCategoryID != null)
                                        category.ParentProductCategoryID = (Guid)parentCategoryID;
                                }
                            }
                            break;

                        case "COSTCODE":
                            data = CellData("COSTCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.CostCode = data.Item1;
                            }
                            break;

                        case "REVENUECODE":
                            data = CellData("REVENUECODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.RevenueCode = data.Item1;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null)
                            {
                                category.Description = data.Item1;
                            }
                            break;
                        case "QUICKBOOKS/XERORENTACCOUNTINGITEMCODE":
                            data = CellData("QUICKBOOKS/XERORENTACCOUNTINGITEMCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.QuickBooksSyncCode = data.Item1;
                            }
                            break;
                        case "QUICKBOOKS/XERODESCRIPTION":
                            data = CellData("QUICKBOOKS/XERODESCRIPTION", ws, row, importDict);
                            if (data != null)
                            {
                                category.QuickBooksSyncCodeDescription = data.Item1;
                            }
                            break;
                        case "SAGE50RENTREVENUEGLCODE":
                            data = CellData("SAGE50RENTREVENUEGLCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.Sage50RevenueRentCode = data.Item1;
                            }
                            break;
                        case "SAGE50SALESREVENUEGLCODE":
                            data = CellData("SAGE50SALESREVENUEGLCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.Sage50RevenueSellCode = data.Item1;
                            }
                            break;
                        case "SAGE50RENTCOSTGLCODE":
                            data = CellData("SAGE50RENTCOSTGLCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.Sage50CostRentCode = data.Item1;
                            }
                            break;
                        case "SAGE50SALESCOSTGLCODE":
                            data = CellData("SAGE50SALESCOSTGLCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.Sage50CostSellCode = data.Item1;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (category.IsSavable)
                {
                    importObject = category;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in category.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportProductCatalog(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
                   List<string> seenProducts)
        {
            bool importFailure = false;
            Product product = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            bool numberProvided = false;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "PARTNUMBER":
                            data = CellData("PARTNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (seenProducts.Contains(data.Item1.ToUpper()))
                                {
                                    {
                                        WriteCell(row, resultCol, ws, "Duplicate Part Number on sheet; can't import this product. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                }
                                else
                                {
                                    product = Product.GetProduct(data.Item1, ProductType.Product);
                                    if (product != null && product.PartNumber.ToUpper() == data.Item1.ToUpper())
                                    {
                                        WriteCell(row, resultCol, ws, "Product exists; can't import this product. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        product = Product.NewProduct(ProductType.Product);
                                        product.PartNumber = data.Item1;
                                        seenProducts.Add(data.Item1.ToUpper());
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Part Number is required;  can't import this product. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null)
                            {
                                product.Description = data.Item1;
                            }
                            break;

                        case "WEIGHT":
                            data = CellData("WEIGHT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.Weight = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Weight is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "MODELNO":
                            data = CellData("MODELNO", ws, row, importDict);
                            if (data != null)
                            {
                                product.ModelNumber = data.Item1;
                            }
                            break;

                        case "CATEGORY":
                            data = CellData("CATEGORY", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                ProductCategory category = ProductCategory.GetProductCategory(data.Item1, ProductType.Product);
                                if (category == null || category.Name != data.Item1)
                                {
                                    WriteCell(row, resultCol, ws, "Product Category isn't recognized; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    product.ProductCategoryID = category.ProductCategoryID;
                                }
                            }
                            break;
                        case "LIST":
                            data = CellData("LIST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultList = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "List is not a decimal number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;
                        case "COST":
                            data = CellData("COST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultCost = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Cost is not a decimal number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;
                        case "SERVICE":
                            data = CellData("SERVICE", ws, row, importDict);
                            BoolResult result = CellBool(data, true, true);

                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Service must be TRUE, FALSE or <empty>. Can't validate this row. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                product.DefaultServiced = result == BoolResult.True ? true : false; // default false
                            }
                            break;

                        case "LENGTH":
                            data = CellData("LENGTH", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.Length = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Length is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "WIDTH":
                            data = CellData("WIDTH", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.Width = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Width is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "HEIGHT":
                            data = CellData("HEIGHT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.Height = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Height is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "CUSTOM1":
                            data = CellData("CUSTOM1", ws, row, importDict);
                            if (data != null)
                            {
                                product.Custom1 = data.Item1;
                            }
                            break;

                        case "CUSTOM2":
                            data = CellData("CUSTOM2", ws, row, importDict);
                            if (data != null)
                            {
                                product.Custom2 = data.Item1;
                            }
                            break;

                        case "CUSTOM3":
                            data = CellData("CUSTOM3", ws, row, importDict);
                            if (data != null)
                            {
                                product.Custom3 = data.Item1;
                            }
                            break;
                        case "MANUFACTURER":
                            data = CellData("MANUFACTURER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                Manufacturer mfg = Manufacturer.GetManufacturer(data.Item1);
                                if (mfg == null || mfg.Name != data.Item1)
                                {
                                    WriteCell(row, resultCol, ws, "Manufacturer is not recognized; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    product.ManufacturerID = mfg.ManufacturerID;
                                }
                            }
                            break;
                        case "MANUFACTURER'SPARTNUMBER":
                            data = CellData("MANUFACTURER'SPARTNUMBER", ws, row, importDict);
                            if (data != null)
                            {
                                product.ManufacturerPartNumber = data.Item1;
                            }
                            break;

                        case "ROP":
                            data = CellData("ROP", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultReOrderPoint = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "ROP is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;
                        case "ROQ":
                            data = CellData("ROQ", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultReorderQuantity = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "ROQ is not a number; can't import this product. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (product.IsSavable)
                {
                    importObject = product;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in product.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportConsumablesCategory(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            ProductCategory category = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            bool numberProvided = false;
            Guid? parentCategoryID = null;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "PARENTNAME":
                            data = CellData("PARENTNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                // First see if the parent is imported previously in the sheet (not yet in the database)
                                try
                                {
                                    // Search importedObjects for the parent  Product Category that matches and is validated previously
                                    var id = importedObjects.Where(x => ((ProductCategory)x.Item1).Name == data.Item1).Select(x => ((ProductCategory)x.Item1).ProductCategoryID).Single();
                                    parentCategoryID = id;
                                }
                                catch
                                {
                                    // See if the databae has the Parent Product Category
                                    ProductCategory parent = ProductCategory.GetProductCategory(data.Item1, ProductType.Product);
                                    if (category == null || category.Name.ToUpper() != data.Item1.ToUpper())
                                    {

                                        WriteCell(row, resultCol, ws, "Parent category doesn't exist in the database.  Cant validate this row. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        parentCategoryID = category.ProductCategoryID;
                                    }
                                }
                            }
                            break;
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                category = ProductCategory.GetProductCategory(data.Item1, ProductType.Product);
                                if (category != null && category.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Consumables category exists; can't import this category. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    category = ProductCategory.NewProductCategory(ProductType.Consumable);
                                    category.Name = data.Item1;
                                    if (parentCategoryID != null)
                                        category.ParentProductCategoryID = (Guid)parentCategoryID;
                                }
                            }
                            break;

                        case "COSTCODE":
                            data = CellData("COSTCODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.CostCode = data.Item1;
                            }
                            break;

                        case "REVENUECODE":
                            data = CellData("REVENUECODE", ws, row, importDict);
                            if (data != null)
                            {
                                category.RevenueCode = data.Item1;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null)
                            {
                                category.Description = data.Item1;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (category.IsSavable)
                {
                    importObject = category;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in category.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportConsumablesCatalog(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
                   List<string> seenConsumables)
        {
            bool importFailure = false;
            Product product = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;

            bool breakLoop = false;
            bool numberProvided = false;


            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "PARTNUMBER":
                            data = CellData("PARTNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (seenConsumables.Contains(data.Item1.ToUpper()))
                                {
                                    WriteCell(row, resultCol, ws, "Part Number can't be duplicated on this sheet. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    seenConsumables.Add(data.Item1.ToUpper());
                                    product = Product.GetProduct(data.Item1, ProductType.Consumable);
                                    if (product != null && product.PartNumber.ToUpper() == data.Item1.ToUpper())
                                    {
                                        WriteCell(row, resultCol, ws, "Consumable exists; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        product = Product.NewProduct(ProductType.Consumable);
                                        product.PartNumber = data.Item1;
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Part Number is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null)
                            {
                                product.Description = data.Item1;
                            }
                            break;

                        case "WEIGHT":
                            data = CellData("WEIGHT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.Weight = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Weight is not a number; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "MODELNO":

                        case "CATEGORY":
                            data = CellData("CATEGORY", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                ProductCategory category = ProductCategory.GetProductCategory(data.Item1, ProductType.Consumable);
                                if (category == null || category.Name != data.Item1)
                                {
                                    WriteCell(row, resultCol, ws, "Consumables Category isn't recognized; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    product.ProductCategoryID = category.ProductCategoryID;
                                }
                            }
                            break;


                        case "COST":
                            data = CellData("COST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultCost = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Cost is not a decimal number; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "MANUFACTURER":
                            data = CellData("MANUFACTURER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                Manufacturer mfg = Manufacturer.GetManufacturer(data.Item1);
                                if (mfg == null || mfg.Name != data.Item1)
                                {
                                    WriteCell(row, resultCol, ws, "Manufacturer is not recognized; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    product.ManufacturerID = mfg.ManufacturerID;
                                }
                            }
                            break;
                        case "MANUFACTURER'SPARTNUMBER":
                            data = CellData("MANUFACTURER'SPARTNUMBER", ws, row, importDict);
                            if (data != null)
                            {
                                product.ManufacturerPartNumber = data.Item1;
                            }
                            break;

                        case "ROP":
                            data = CellData("ROP", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultReOrderPoint = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "ROP is not a number; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;
                        case "ROQ":
                            data = CellData("ROQ", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product.DefaultReorderQuantity = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "ROQ is not a number; can't import this consumable. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (product.IsSavable)
                {
                    importObject = product;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in product.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportRentAdjustment(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol,
                   Dictionary<string, RentalPeriodAdjustment> adjustmentDic,
                   Dictionary<string, List<int>> adjustmentGoodRule,
                   List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            RentalPeriodAdjustment adjust = null;
            RentalPeriodAdjustmentRule rule = null;
            bool breakLoop = false;
            bool numberProvided = false;
            bool adjustCreated = false;
            string adjName = "";

            Tuple<string, ImportConfig.ImportColunnType, int> data;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                adjName = data.Item1.ToUpper();
                                if (adjustmentDic.ContainsKey(adjName))
                                {
                                    adjust = adjustmentDic[adjName];
                                }
                                else
                                {
                                    adjust = RentalPeriodAdjustment.GetRentalPeriodAdjustment(adjName);
                                    if (adjust == null || adjust.Name.ToUpper() != adjName)
                                    {
                                        adjust = RentalPeriodAdjustment.NewRentalPeriodAdjustment();
                                        adjust.Name = data.Item1;
                                        adjust.IsActive = true;  // default
                                        adjustCreated = true;
                                    }
                                    // Note adjust is added to adjustmentDic where tuple is added for import
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Name is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE or FALSE. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                adjust.IsActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                adjust.Description = data.Item1;
                            }
                            break;

                        case "RULENAME":
                            data = CellData("RULENAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                if (adjust.Rules.Where(x => x.Name == data.Item1).Any())
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Rent Period Adjustment Rule Name {0} exist in the database.  Cant validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }

                                rule = RentalPeriodAdjustmentRule.NewRentalPeriodAdjustmentRuleChild();
                                rule.Name = data.Item1;
                                rule.RentalPeriodAdjustmentID = adjust.RentalPeriodAdjustmentID;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("Rule Name is required.  Cant validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "DATEORDAYS":
                            data = CellData("DATEORDAYS", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string cellStr = (data.Item1.ToUpper());
                                if (cellStr == "DATE" || cellStr == "DAYS")
                                {
                                    rule.RangeType = cellStr == "DATE" ? RuleRangeTypes.DateRange : RuleRangeTypes.DayRange;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Date or Days must equal 'DATE' OR 'DAYS'. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                            }
                            break;

                        case "START":
                            data = CellData("START", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (rule.RangeType == RuleRangeTypes.DateRange)
                                {
                                    try
                                    {
                                        rule.StartDate = SmartDate.Parse(data.Item1);
                                    }
                                    catch
                                    {
                                        WriteCell(row, resultCol, ws, string.Format("Start must be a date.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        rule.StartDay = int.Parse(data.Item1);
                                    }
                                    catch
                                    {
                                        WriteCell(row, resultCol, ws, string.Format("Start must be an integer.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("Start is required.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "END":
                            data = CellData("END", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (rule.RangeType == RuleRangeTypes.DateRange)
                                {
                                    try
                                    {
                                        rule.EndDate = SmartDate.Parse(data.Item1);
                                    }
                                    catch
                                    {
                                        WriteCell(row, resultCol, ws, string.Format("End must be a date.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        rule.EndDay = int.Parse(data.Item1);
                                    }
                                    catch
                                    {
                                        WriteCell(row, resultCol, ws, string.Format("End must be an integer.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("End is required.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "DISCOUNTORMARKUP":
                            data = CellData("DISCOUNTORMARKUP", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string cellStr = (data.Item1.ToUpper());
                                if (cellStr == "DISCOUNT" || cellStr == "MARKUP")
                                {
                                    rule.MarkupType = cellStr == "DISCOUNT" ? RuleMarkupTypes.Discount : RuleMarkupTypes.Markup;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Discount or Markup must be 'Discount' or 'Markup'. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("Discount or Markup is required.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "VALUE":
                            data = CellData("VALUE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    rule.Factor = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Value must be a number.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("Discount or Markup is required.  Can't validate this row ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                adjust.Rules.Add(rule);

                object importObject = null;

                if (adjust.IsSavable && rule.IsSavable && !adjustmentDic.ContainsKey(adjName))
                {
                    importObject = adjust;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided,
                            adjustCreated, row, resultCol, adjustmentGoodRule);

                    importedObjects.Add(tuple);
                    adjustmentDic.Add(adjName, adjust);
                    List<int> tempRows = new List<int>();
                    tempRows.Add(row);

                    adjustmentGoodRule.Add(adjName, tempRows);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else if (adjust.IsSavable && rule.IsSavable)
                {
                    // Add this rule to the existing adjustment from a previous row
                    Tuple<object, object, object, int, int, object> tuple = importedObjects.Where(x => ((RentalPeriodAdjustment)x.Item1).Name.ToUpper() == adjName).Single();
                    ((RentalPeriodAdjustment)(tuple.Item1)).Rules.Add(rule);

                    adjustmentGoodRule[adjName].Add(row);

                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in adjust.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Adjustment has Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }

                    foreach (var brokenRule in rule.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Adjustment's Rule has Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportRateProfile(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
                   ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            RateProfile profile = null;

            RateProfileProduct product = null;

            bool breakLoop = false;
            bool numberProvided = false;
            Tuple<string, ImportConfig.ImportColunnType, int> data;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string name = data.Item1.ToUpper();

                                bool found = false;
                                foreach (var o in importedObjects)
                                {
                                    String itemName = ((RateProfile)o.Item1).Name;
                                    if (name == itemName.ToUpper())
                                    {
                                        profile = (RateProfile)o.Item1;
                                        found = true;
                                    }
                                }

                                if (!found)
                                {
                                    profile = RateProfile.NewRateProfile();
                                    profile.Name = data.Item1;
                                    profile.IsActive = true;  // default
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Name is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE, FALSE, YES or NO. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                profile.IsActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                profile.Description = data.Item1;
                            }
                            break;

                        case "LABELFORHEADER":
                            data = CellData("LABELFORHEADER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                profile.RateLabel = data.Item1;
                            }
                            else
                            {
                                profile.RateLabel = "28 Day Rate";
                            }
                            break;
                        case "RENTALDAYS":
                            data = CellData("RENTALDAYS", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    profile.Days = short.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Rate Days is not an integer. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            else
                            {
                                profile.Days = 28;
                            }
                            break;

                        case "PARTNUMBER":
                            data = CellData("PARTNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    product = (profile.RateProfileProducts.Where(x => x.PartNumber.ToUpper() == data.Item1.ToUpper())
                                        .Union(profile.RateProfileConsumables.Where(x => x.PartNumber.ToUpper() == data.Item1.ToUpper()))).Single();

                                    // TODO: Remove this stuff
                                    // -------------------------------
                                    // profile.RateProfileProducts
                                    // profile.RateProfileConsumables
                                    // -------------------------------
                                    // product = profile.RateProfileProducts.Where(x => x.PartNumber.ToUpper() == data.Item1.ToUpper()).Single();
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Part Number is not recognized. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }

                                if (!breakLoop && product.RateProfileProductID == Guid.Empty)
                                {
                                    product.RateProfileProductID = Guid.NewGuid();
                                    product.RateProfileID = profile.RateProfileID;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("Part Number is required. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "RATE":
                            data = CellData("RATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.RentRate = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Rate is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "RENTQTY":
                            data = CellData("RENTQTY", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.ProratedQuantity = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Rent Qty is not a number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "JOBCOST":
                            data = CellData("JOBCOST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.JobCost = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Job Cost is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "DELIVERYCHARGE":
                            data = CellData("DELIVERYCHARGE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.DeliveryCharge = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Delivery Charge is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "RETURNCREDIT":

                            data = CellData("RETURNCREDIT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.ReturnCredit = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Return Credit is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SELLPRICE":
                            data = CellData("SELLPRICE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    if (product.ProductType == ProductType.Consumable)
                                    {
                                        product.RentRate = decimal.Parse(data.Item1);
                                    }
                                    else
                                    {
                                        product.SellPrice = decimal.Parse(data.Item1);
                                    }
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Sell Price is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "CHARGE/UNIT":
                            data = CellData("CHARGE/UNIT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    decimal TOBEDETERMINED = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Charge/Unit is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "REPLACEMENTCOST":
                            data = CellData("REPLACEMENTCOST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.ReplacementCost = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Replacement Cost is not a decimal number. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "MIN.DAYS":
                            data = CellData("MIN.DAYS", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                try
                                {
                                    product.MinimumDayOverride = short.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Min. Days is not an integer. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "RENTADJUSTMENT":
                            data = CellData("RENTADJUSTMENT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                RentalPeriodAdjustment adjustment = RentalPeriodAdjustment.GetRentalPeriodAdjustment(data.Item1);
                                if (adjustment != null && adjustment.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    product.RentalPeriodAdjustmentID = adjustment.RentalPeriodAdjustmentID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Rental Period Adjustment not recognized. Can't validate this row. ", data.Item1), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;


                if (profile.IsSavable && product.IsSavable)
                {
                    importObject = profile;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in profile.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Rate Profile has Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }

                    foreach (var brokenRule in product.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Rate Profile Product has Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportCustomer(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            List<string> namesSeen)
        {
            bool importFailure = false;
            BusinessPartner customer = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            bool numberProvided = false;
            Address businessAddress = null;
            Address billingAddress = null;
            Address shippingAddress = null;
            State state = null;
            County county = null;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (namesSeen.Contains(data.Item1.ToUpper()))
                                {
                                    WriteCell(row, resultCol, ws, "Duplicate Customer Name is not allowed; can't import this customer. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    namesSeen.Add(data.Item1.ToUpper());
                                    customer = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                    if (customer != null && customer.Name.ToUpper() == data.Item1.ToUpper())
                                    {
                                        WriteCell(row, resultCol, ws, "Customer exists; can't import this customer. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        customer = BusinessPartner.NewBusinessPartner(PartnerTypes.Customer);
                                        customer.Name = data.Item1;
                                    }
                                }
                            }
                            break;
                        case "NUMBER":
                            data = CellData("NUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                customer.PartnerNumber = data.Item1;
                                numberProvided = true;
                            }
                            break;

                        case "ACCOUNTINGID":
                            data = CellData("ACCOUNTINGID", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                customer.AccountingID = data.Item1;
                            }
                            break;

                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE or FALSE. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                customer.IsActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "EMAIL":
                            data = CellData("EMAIL", ws, row, importDict);
                            if (data != null)
                            {
                                customer.EmailAddress = data.Item1;
                            }
                            break;
                        case "WEBSITE":
                            data = CellData("WEBSITE", ws, row, importDict);
                            if (data != null)
                            {
                                customer.WebAddress = data.Item1;
                            }
                            break;
                        case "BUSINESSPHONE":
                            data = CellData("BUSINESSPHONE", ws, row, importDict);
                            if (data != null)
                            {
                                customer.PhoneNumber = data.Item1;
                            }
                            break;

                        case "FAX#":
                            data = CellData("FAX#", ws, row, importDict);
                            if (data != null)
                            {
                                customer.FaxNumber = data.Item1;
                            }
                            break;

                        case "BUSINESSSTREET1":
                            data = CellData("BUSINESSSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                businessAddress.Street = data.Item1;
                            }
                            break;

                        case "BUSINESSSTREET2":
                            data = CellData("BUSINESSSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                businessAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BUSINESSCITY":
                            data = CellData("BUSINESSCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                businessAddress.City = data.Item1;
                            }
                            break;

                        case "BUSINESSSTATE":
                            data = CellData("BUSINESSSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            data = CellData("BUSINESSCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.CountyID = county.CountyID;
                                }
                            }
                            break;



                        case "BUSINESSZIP/POSTALCODE":
                            data = CellData("BUSINESSZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                businessAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            data = CellData("BUSINESSCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                businessAddress.Country = data.Item1;
                            }
                            break;
                        case "BUSINESSLATITUDE":
                            data = CellData("BUSINESSLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }
                                try
                                {
                                    businessAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BUSINESSLONGITUDE":
                            data = CellData("BUSINESSLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            data = CellData("BUSINESSELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(customer, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGSTREET1":
                            data = CellData("BILLINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }

                                billingAddress.Street = data.Item1;
                            }
                            break;

                        case "BILLINGSTREET2":
                            data = CellData("BILLINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }
                                billingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BILLINGCITY":
                            data = CellData("BILLINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }
                                billingAddress.City = data.Item1;
                            }
                            break;

                        case "BILLINGSTATE":
                            data = CellData("BILLINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            data = CellData("BILLINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            data = CellData("BILLINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }
                                billingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            data = CellData("BILLINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }
                                billingAddress.Country = data.Item1;
                            }
                            break;
                        case "BILLINGLATITUDE":
                            data = CellData("BILLINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }
                                try
                                {
                                    billingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGLONGITUDE":
                            data = CellData("BILLINGLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            data = CellData("BILLINGELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(customer, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGSTREET1":
                            data = CellData("SHIPPINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }

                                shippingAddress.Street = data.Item1;
                            }
                            break;

                        case "SHIPPINGSTREET2":
                            data = CellData("SHIPPINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }
                                shippingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "SHIPPINGCITY":
                            data = CellData("SHIPPINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }
                                shippingAddress.City = data.Item1;
                            }
                            break;


                        case "SHIPPINGSTATE":
                            data = CellData("SHIPPINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            data = CellData("SHIPPINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;

                        case "SHIPPINGZIP/POSTALCODE":
                            data = CellData("SHIPPINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }
                                shippingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            data = CellData("SHIPPINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }
                                shippingAddress.Country = data.Item1;
                            }
                            break;
                        case "SHIPPINGLATITUDE":
                            data = CellData("SHIPPINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }
                                try
                                {
                                    shippingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGLONGITUDE":
                            data = CellData("SHIPPINGLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "SHIPPINGELEVATION":
                            data = CellData("SHIPPINGELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(customer, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (customer.IsSavable)
                {
                    importObject = customer;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in customer.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportCustomerContact(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            AvontusUser user = AvontusUser.NewUser(UserTypes.Customer);

            int userNameColumn = config.ImportColumns.Where(x => NormStr(x.ColumnHeading) == "USERNAME").First().Column;
            BusinessPartner customer = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    //importFailure = false;
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "CUSTOMERNAME":
                            data = CellData("CUSTOMERNAME", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                customer = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                if (customer != null && customer.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    user.RelatedID = customer.BusinessPartnerID;
                                    user.RoleID = RoleList.GetRoles(false, true).GetRoleByType(AvontusRoleType.Customer).RoleID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Customer name not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Customer name required, not provided. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "FIRST":
                            data = CellData("FIRST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.FirstName = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "First name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "LAST":
                            data = CellData("LAST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.LastName = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Last name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "TITLE":
                            data = CellData("TITLE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Title = data.Item1;
                            }
                            break;

                        case "EMAIL":
                            data = CellData("EMAIL", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.EmailAddress = data.Item1;
                            }
                            break;

                        case "PHONE":
                            data = CellData("PHONE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.PhoneNumber = data.Item1;
                            }
                            break;
                        case "FAX":
                            data = CellData("FAX", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.FaxNumber = data.Item1;
                            }
                            break;

                        case "CELL":
                            data = CellData("CELL", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.CellNumber = data.Item1;
                            }
                            break;

                        case "NOTES":
                            data = CellData("NOTES", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Notes = data.Item1;
                            }
                            break;


                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE or FALSE. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                user.IsActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "USERNAME":
                            data = CellData("USERNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                AvontusUser existingUser = AvontusUser.GetUser(data.Item1);
                                if (existingUser != null && existingUser.Username.ToUpper() == data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "User Name is already in use. Must be unique. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                user.Username = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "User Name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;
                        case "PASSWORD":
                            data = CellData("PASSWORD", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Password = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Password is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;


                        default:
                            break;
                    }
                }
            }
            if (!importFailure)
            {
                object importObject = null;
                if (user.IsSavable)
                {
                    importObject = user;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, null, null, row, resultCol, null);
                    WriteCell(row, resultCol, ws, Validated);
                    importedObjects.Add(tuple);
                }
                else
                {
                    foreach (var brokenRule in user.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportOrder(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, Dictionary<string, OrderInfo> orderDic,
            Dictionary<string, OrderInfo> workOrderDic, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            BusinessPartner customer = null;

            Order order = null;
            bool orderCreated = false;


            Order workOrder = null;
            bool workOrderCreated = false;

            Tuple<string, ImportConfig.ImportColunnType, int> data;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "CUSTOMERNAME":
                            data = CellData("CUSTOMERNAME", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                customer = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                if (customer == null || customer.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Customer name not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Customer Name is required and not populated", useColor: true, append: true, color: warningColor);
                                importFailure = true;

                            }
                            break;

                        case "ORDERNUMBER":
                            data = CellData("ORDERNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (orderDic.ContainsKey(data.Item1))
                                {
                                    order = orderDic[data.Item1].Order;
                                }
                                else
                                {
                                    order = Order.GetOrder(data.Item1);
                                    if (customer != null)
                                    {
                                        if (order.Number != data.Item1)
                                        {
                                            order = Order.NewPurchaseOrder(customer.BusinessPartnerID);
                                            order.Number = data.Item1;
                                            // orderDic.Add(data.Item1, new OrderInfo(order, false, false));
                                            orderCreated = true;
                                        }
                                        else
                                        {
                                            orderDic.Add(data.Item1, new OrderInfo(order, false, true));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Order Number is required and not populated", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;


                        case "WORKORDERNUMBER":
                            data = CellData("WORKORDERNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string combinedOrderNumber = order.Number + "|" + data.Item1;

                                if (workOrderDic.ContainsKey(combinedOrderNumber))
                                {
                                    workOrder = workOrderDic[combinedOrderNumber].Order;
                                }
                                else
                                {
                                    if (order.Number != "")
                                    {
                                        workOrder = Order.GetWorkOrder(order.Number, data.Item1);
                                        if (workOrder.Number != data.Item1 || workOrder.ParentOrderID != order.OrderID)
                                        {
                                            workOrder = Order.NewWorkOrder(customer.BusinessPartnerID, order.OrderID);
                                            workOrder.Number = data.Item1;
                                            workOrderCreated = true;
                                        }
                                        else
                                        {
                                            workOrderDic.Add(data.Item1, new OrderInfo(workOrder, true, true));
                                        }
                                    }
                                }
                            }
                            break;

                        case "ORDERISACTIVE":
                            data = CellData("ORDERISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Order IsActive must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                order.IsActive = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "ORDERDESCRIPTION":
                            data = CellData("ORDERDESCRIPTION", ws, row, importDict);
                            order.Description = data.Item1 ?? "";
                            break;

                        case "ORDERAPPROVERUSERNAME":
                            data = CellData("ORDERAPPROVERUSERNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                AvontusUser user = AvontusUser.GetUser(data.Item1);
                                if (user.Username.ToUpper() == data.Item1.ToUpper())
                                {
                                    order.ApproverID = user.UserID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Order Approver User Name is not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    break;
                                }
                            }
                            break;

                        case "ORDERREVIEWERUSERNAME":
                            data = CellData("ORDERREVIEWERUSERNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                AvontusUser user = AvontusUser.GetUser(data.Item1);
                                if (user.Username.ToUpper() == data.Item1.ToUpper())
                                {
                                    order.ReviewerID = user.UserID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Order Reviewer User Name is not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    break;
                                }
                            }
                            break;

                        case "ORDERAMOUNT":
                            data = CellData("ORDERAMOUNT", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                try
                                {
                                    order.ApprovedAmount = decimal.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "Order Amount is not a valid number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;

                                }
                            }
                            break;

                        case "WORKORDERISACTIVE":
                            data = CellData("WORKORDERISACTIVE", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Work Order IsActive must be TRUE, FALSE, YES, NO or blank (when there's no work order).  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                if (workOrder != null)
                                {
                                    workOrder.IsActive = result == BoolResult.False ? false : true;
                                }
                            }
                            break;

                        case "WORKORDERDESCRIPTION":
                            if (workOrder != null)
                            {

                                data = CellData("WORKORDERDESCRIPTION", ws, row, importDict);
                                workOrder.Description = data.Item1 ?? "";
                            }
                            break;

                        case "WORKORDERAPPROVERUSERNAME":
                            if (workOrder != null)
                            {
                                data = CellData("WORKORDERAPPROVERUSERNAME", ws, row, importDict);
                                if (data != null && data.Item1 != "")
                                {
                                    AvontusUser user = AvontusUser.GetUser(data.Item1);
                                    if (user.Username.ToUpper() == data.Item1.ToUpper())
                                    {
                                        workOrder.ApproverID = user.UserID;
                                    }
                                    else
                                    {
                                        WriteCell(row, resultCol, ws, "Work Order Approver User Name is not recognized. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        break;
                                    }
                                }
                            }
                            break;

                        case "WORKORDERREVIEWERUSERNAME":
                            if (workOrder != null)
                            {
                                data = CellData("WORKORDERREVIEWERUSERNAME", ws, row, importDict);
                                if (data != null && data.Item1 != "")
                                {
                                    AvontusUser user = AvontusUser.GetUser(data.Item1);
                                    if (user.Username.ToUpper() == data.Item1.ToUpper())
                                    {
                                        workOrder.ReviewerID = user.UserID;
                                    }
                                    else
                                    {
                                        WriteCell(row, resultCol, ws, "Work Order Reviewer User Name is not recognized. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        break;
                                    }
                                }
                            }
                            break;
                        case "WORKORDERAMOUNT":
                            if (workOrder != null)
                            {
                                data = CellData("WORKORDERAMOUNT", ws, row, importDict);
                                if (data != null && data.Item1 != "")
                                {
                                    try
                                    {
                                        workOrder.ApprovedAmount = decimal.Parse(data.Item1);
                                    }
                                    catch
                                    {
                                        WriteCell(row, resultCol, ws, "Work Order Amount is not a valid number. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                    }
                                }
                            }
                            break;
                        default:
                            break;

                    }
                }
            }

            if (!importFailure)
            {
                if (orderCreated && !order.IsSavable)
                {
                    foreach (var brokenRule in order.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Order Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
                else if (workOrderCreated && !workOrder.IsSavable)
                {
                    foreach (var brokenRule in workOrder.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Work Order Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
                else if (!orderCreated && !workOrderCreated)
                {
                    if (workOrder == null)
                    {
                        WriteCell(row, resultCol, ws, "Order exists in database; can't import. ", useColor: true, append: true, color: warningColor);
                    }
                    else
                    {
                        WriteCell(row, resultCol, ws, "Order and work order exist in database; can't import. ", useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
                else
                {
                    if (orderCreated && workOrderCreated)
                    {
                        Tuple<object, object, object, int, int, object>
                            tuple = new Tuple<object, object, object, int, int, object>(order, workOrder, null, row, resultCol, null);
                        WriteCell(row, resultCol, ws, Validated);
                        importedObjects.Add(tuple);
                        orderDic.Add(order.Number, new OrderInfo(order, false, false));
                        workOrderDic.Add(order.Number + "|" + workOrder.Number, new OrderInfo(workOrder, true, false));
                    }
                    else if (workOrderCreated)
                    {
                        Tuple<object, object, object, int, int, object>
                            tuple = new Tuple<object, object, object, int, int, object>(null, workOrder, null, row, resultCol, null);
                        WriteCell(row, resultCol, ws, Validated);
                        importedObjects.Add(tuple);
                        workOrderDic.Add(order.Number + "|" + workOrder.Number, new OrderInfo(workOrder, true, false));
                    }
                    else if (orderCreated)
                    {
                        Tuple<object, object, object, int, int, object>
                            tuple = new Tuple<object, object, object, int, int, object>(order, null, null, row, resultCol, null);
                        WriteCell(row, resultCol, ws, Validated);
                        importedObjects.Add(tuple);
                        orderDic.Add(order.Number, new OrderInfo(order, false, false));
                    }
                }
            }

            return importFailure;
        }

        private bool ImportVendor(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
           ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            BusinessPartner vendor = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            bool numberProvided = false;
            Address businessAddress = null;
            Address billingAddress = null;
            Address shippingAddress = null;
            State state = null;
            County county = null;

            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                vendor = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                if (vendor != null && vendor.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Vendor exists; can't import this vendor. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    vendor = BusinessPartner.NewBusinessPartner(PartnerTypes.Vendor);
                                    vendor.Name = data.Item1;
                                }
                            }
                            break;
                        case "NUMBER":
                            data = CellData("NUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                vendor.PartnerNumber = data.Item1;
                                numberProvided = true;
                            }
                            break;

                        case "ACCOUNTINGID":
                            data = CellData("ACCOUNTINGID", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                vendor.AccountingID = data.Item1;
                            }
                            break;
                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be TRUE or FALSE. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                vendor.IsActive = result == BoolResult.False ? false : true;
                            }
                            break;

                        case "EMAIL":
                            data = CellData("EMAIL", ws, row, importDict);
                            if (data != null)
                            {
                                vendor.EmailAddress = data.Item1;
                            }
                            break;
                        case "WEBSITE":
                            data = CellData("WEBSITE", ws, row, importDict);
                            if (data != null)
                            {
                                vendor.WebAddress = data.Item1;
                            }
                            break;
                        case "BUSINESSPHONE":
                            data = CellData("BUSINESSPHONE", ws, row, importDict);
                            if (data != null)
                            {
                                vendor.PhoneNumber = data.Item1;
                            }
                            break;

                        case "FAX#":
                            data = CellData("FAX#", ws, row, importDict);
                            if (data != null)
                            {
                                vendor.FaxNumber = data.Item1;
                            }
                            break;

                        case "BUSINESSSTREET1":
                            data = CellData("BUSINESSSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                businessAddress.Street = data.Item1;
                            }
                            break;

                        case "BUSINESSSTREET2":
                            data = CellData("BUSINESSSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                businessAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BUSINESSCITY":
                            data = CellData("BUSINESSCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                businessAddress.City = data.Item1;
                            }
                            break;

                        case "BUSINESSSTATE":
                            data = CellData("BUSINESSSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            data = CellData("BUSINESSCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.CountyID = county.CountyID;
                                }
                            }
                            break;

                        case "BUSINESSZIP/POSTALCODE":
                            data = CellData("BUSINESSZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                businessAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            data = CellData("BUSINESSCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                businessAddress.Country = data.Item1;
                            }
                            break;
                        case "BUSINESSLATITUDE":
                            data = CellData("BUSINESSLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }
                                try
                                {
                                    businessAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BUSINESSLONGITUDE":
                            data = CellData("BUSINESSLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            data = CellData("BUSINESSELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetPartnerAddress(vendor, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGSTREET1":
                            data = CellData("BILLINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }

                                billingAddress.Street = data.Item1;
                            }
                            break;

                        case "BILLINGSTREET2":
                            data = CellData("BILLINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }
                                billingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BILLINGCITY":
                            data = CellData("BILLINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }
                                billingAddress.City = data.Item1;
                            }
                            break;

                        case "BILLINGSTATE":
                            data = CellData("BILLINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            data = CellData("BILLINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            data = CellData("BILLINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }
                                billingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            data = CellData("BILLINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }
                                billingAddress.Country = data.Item1;
                            }
                            break;
                        case "BILLINGLATITUDE":
                            data = CellData("BILLINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }
                                try
                                {
                                    billingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGLONGITUDE":
                            data = CellData("BUSINESSLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            data = CellData("BUSINESSELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetPartnerAddress(vendor, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGSTREET1":
                            data = CellData("SHIPPINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }

                                shippingAddress.Street = data.Item1;
                            }
                            break;

                        case "SHIPPINGSTREET2":
                            data = CellData("SHIPPINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }
                                shippingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "SHIPPINGCITY":
                            data = CellData("SHIPPINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }
                                shippingAddress.City = data.Item1;
                            }
                            break;

                        case "SHIPPINGSTATE":
                            data = CellData("SHIPPINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            data = CellData("SHIPPINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "SHIPPINGZIP/POSTALCODE":
                            data = CellData("SHIPPINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }
                                shippingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            data = CellData("SHIPPINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }
                                shippingAddress.Country = data.Item1;
                            }
                            break;
                        case "SHIPPINGLATITUDE":
                            data = CellData("SHIPPINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }
                                try
                                {
                                    shippingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGLONGITUDE":
                            data = CellData("BUSINESSLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "SHIPPINGELEVATION":
                            data = CellData("BUSINESSELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetPartnerAddress(vendor, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject = null;
                if (vendor.IsSavable)
                {
                    importObject = vendor;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, numberProvided, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in vendor.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportVendorContact(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            AvontusUser user = AvontusUser.NewUser(UserTypes.Vendor);

            int userNameColumn = config.ImportColumns.Where(x => NormStr(x.ColumnHeading) == "USERNAME").First().Column;
            BusinessPartner vendor = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    //importFailure = false;
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "VENDORNAME":
                            data = CellData("VENDORNAME", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                vendor = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                if (vendor != null && vendor.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    user.RelatedID = vendor.BusinessPartnerID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Vendor name not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;

                        case "FIRST":
                            data = CellData("FIRST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.FirstName = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "First name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "LAST":
                            data = CellData("LAST", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.LastName = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Last name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;
                        case "TITLE":
                            data = CellData("TITLE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Title = data.Item1;
                            }
                            break;

                        case "EMAIL":
                            data = CellData("EMAIL", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.EmailAddress = data.Item1;
                            }
                            break;
                        case "PHONE":
                            data = CellData("PHONE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.PhoneNumber = data.Item1;
                            }
                            break;
                        case "FAX":
                            data = CellData("FAX", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.FaxNumber = data.Item1;
                            }
                            break;
                        case "CELL":
                            data = CellData("CELL", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.CellNumber = data.Item1;
                            }
                            break;
                        case "NOTES":
                            data = CellData("NOTES", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Notes = data.Item1;
                            }
                            break;

                        case "USERNAME":
                            data = CellData("USERNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                AvontusUser existingUser = AvontusUser.GetUser(data.Item1);
                                if (existingUser != null && existingUser.Username.ToUpper() == data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "User Name is already in use. Must be unique. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                user.Username = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "User Name is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;
                        case "PASSWORD":
                            data = CellData("PASSWORD", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                user.Password = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Password is required and not valid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }
            if (!importFailure)
            {
                object importObject = null;
                if (user.IsSavable)
                {
                    importObject = user;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, null, null, row, resultCol, null);
                    WriteCell(row, resultCol, ws, Validated);
                    // WriteCell(ws, row, userNameColumn + 1, resultCol, "Validated.", validationColor, true);
                    importedObjects.Add(tuple);
                }
                else
                {
                    foreach (var brokenRule in user.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportTaxRate(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            HashSet<string> foundNames)
        {
            bool importFailure = false;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            BoolResult result = BoolResult.Invalid;
            TaxRate taxRate = null;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);

                            if (data == null || data.Item1 == "")
                            {
                                WriteCell(row, resultCol, ws, "Name is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                string tempName = data.Item1.ToUpper();
                                taxRate = TaxRate.GetTaxRate(tempName);
                                if (taxRate != null && taxRate.Name.ToUpper() == tempName)
                                {
                                    WriteCell(row, resultCol, ws, "The Tax Rate exists in the datase. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else if (foundNames.Contains(tempName))
                                {
                                    WriteCell(row, resultCol, ws, "Name is a duplicate on this sheet. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    taxRate = TaxRate.NewTaxRate();
                                    taxRate.Name = data.Item1;
                                    foundNames.Add(tempName);
                                }
                            }
                            break;
                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            result = CellBool(data.Item1, true, true);

                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Active must be True, False, Yes, No or blank (defaulting True). ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                taxRate.IsActive = (result == BoolResult.True || result == BoolResult.Empty);
                            }
                            break;
                        case "REFID":
                            data = CellData("REFID", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                taxRate.RefID = data.Item1;
                            }
                            break;
                        case "TAXAGENCY":
                            data = CellData("TAXAGENCY", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                taxRate.TaxAgency = data.Item1;
                            }
                            break;
                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                taxRate.Description = data.Item1;
                            }
                            break;

                        case "RATE":
                            data = CellData("RATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                bool isNumber = double.TryParse(data.Item1, out double tempDouble);
                                if (!isNumber)
                                {
                                    WriteCell(row, resultCol, ws, "Rate is not a number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    taxRate.Rate = tempDouble;
                                }
                            }
                            else
                            {
                                taxRate.Rate = 0.0;  // default  value
                            }
                            break;

                        default:
                            break;
                    }
                }
            }
            if (!importFailure)
            {
                object importObject = null;
                if (taxRate.IsSavable)
                {
                    importObject = taxRate;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, null, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in taxRate.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportJobsite(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            HashSet<string> jobNamesInSheet)
        {
            bool importFailure = false;
            bool trackRequests = GlobalClientConfig.OptionsSettings.EnableRequestPortal;
            Guid parentID = Guid.Empty;
            string parentName = "";
            Guid customerID = Guid.Empty;
            StockingLocation jobsite = null;
            Order order = null;
            BusinessPartner customer = null;
            Tuple<string, ImportConfig.ImportColunnType, int> data;
            Address billingAddress = null;
            Address shippingAddress = null;
            Address businessAddress = null;
            State state = null;
            County county = null;
            BoolResult result = BoolResult.Invalid;


            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "PARENTLOCATION":
                            data = CellData("PARENTLOCATION", ws, row, importDict);

                            if (data == null && data.Item1 == "")
                            {
                                WriteCell(row, resultCol, ws, "Parent Location must be identified. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }

                            TradingPartner parent = TradingPartner.GetTradingPartner(data.Item1, PartnerTypes.BranchOffice);
                            if (parent == null || parent.Name.ToUpper() != data.Item1.ToUpper())
                            {
                                parent = TradingPartner.GetTradingPartner(data.Item1, PartnerTypes.LaydownYard);

                                if (parent == null || parent.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    parent = TradingPartner.GetTradingPartner(data.Item1, PartnerTypes.StagingArea);
                                }
                            }

                            if (parent == null || parent.Name.ToUpper() != data.Item1.ToUpper())
                            {
                                parentID = Guid.NewGuid();
                                parentName = data.Item1.ToUpper();
                                if (!jobNamesInSheet.Contains(parentName))
                                {
                                    WriteCell(row, resultCol, ws, "Parent Location is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            else
                            {
                                parentID = parent.TradingPartnerID;
                            }
                            break;
                        case "CUSTOMERNAME":
                            data = CellData("CUSTOMERNAME", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                customer = BusinessPartner.GetBusinessPartnerByName(data.Item1);
                                if (customer != null && customer.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    customerID = customer.BusinessPartnerID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Customer name not recognized. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;

                        case "NAME":
                            data = CellData("NAME", ws, row, importDict);

                            if (data == null || data.Item1 == "")
                            {
                                WriteCell(row, resultCol, ws, "Jobsite name is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else
                            {
                                jobsite = StockingLocation.GetStockingLocation(data.Item1, false);
                                if (jobsite == null || jobsite.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    jobsite = StockingLocation.NewStockingLocation(PartnerTypes.JobSite, parentID);
                                    jobsite.Name = data.Item1;
                                    jobNamesInSheet.Add(data.Item1.ToUpper());
                                    jobsite.BusinessPartnerID = customerID;
                                    jobsite.ParentTradingPartnerID = parentID;
                                    jobsite.TrackRequests = true;  // ??
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Jobsite exists; unable to import. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "NUMBER":
                            data = CellData("NUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite.Number = data.Item1;
                            }
                            break;
                        case "ACCOUNTINGID":
                            data = CellData("ACCOUNTINGID", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite.AccountingID = data.Item1;
                            }
                            break;

                        case "DESCRIPTION":
                            data = CellData("DESCRIPTION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite.Description = data.Item1.Replace('\"', '\'');
                            }
                            break;
                        case "ISACTIVE":
                            data = CellData("ISACTIVE", ws, row, importDict);
                            result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "IsActive is Required must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.IsActive = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "ISBILLABLE":
                            data = CellData("ISBILLABLE", ws, row, importDict);
                            result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Is Billable  must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else if (result == BoolResult.True)
                            {
                                jobsite.IsBillable = true;
                                jobsite.InvoiceOrderGrouping = InvoiceOrderGroupType.PerJob;
                            }
                            else
                            {
                                jobsite.IsBillable = false;
                            }
                            break;

                        case "TRACKACTIVITIES":
                            data = CellData("TRACKACTIVITIES", ws, row, importDict);
                            result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Track Activities must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.UseScaffoldTagActivities = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "TRACKSCAFFOLDS":
                            data = CellData("TRACKSCAFFOLDS", ws, row, importDict);
                            result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Track Scaffolds must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.TrackScaffoldTags = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "REQUESTS":
                            data = CellData("REQUESTS", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.True && !trackRequests)
                            {
                                WriteCell(row, resultCol, ws, "Global option to Track Requests is not checked; unable to import with Requests = TRUE or Yes. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;

                            }
                            else if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Requests must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.TrackRequests = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDACTIVITYLIST1LABEL":
                            data = CellData("SCAFFOLDACTIVITYLIST1LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagActivityList1Label = data.Item1;
                                jobsite.UseScaffoldTagActivityList1 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST1REQUIRED":
                            data = CellData("SCAFFOLDACTIVITYLIST1REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Activity List1 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagActivityList1Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDACTIVITYLIST2LABEL":
                            data = CellData("SCAFFOLDACTIVITYLIST2LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagActivityList2Label = data.Item1;
                                jobsite.UseScaffoldTagActivityList2 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST2REQUIRED":
                            data = CellData("SCAFFOLDACTIVITYLIST2REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Activity List2 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagActivityList2Required = result == BoolResult.True ? true : false;
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST3LABEL":
                            data = CellData("SCAFFOLDACTIVITYLIST3LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagActivityList3Label = data.Item1;
                                jobsite.UseScaffoldTagActivityList3 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST3REQUIRED":
                            data = CellData("SCAFFOLDACTIVITYLIST3REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Activity List3 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagActivityList3Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDACTIVITYTEXTLABEL":
                            data = CellData("SCAFFOLDACTIVITYTEXTLABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagActivityText1Label = data.Item1;
                                jobsite.UseScaffoldTagActivityText1 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDACTIVITYTEXTREQUIRED":
                            data = CellData("SCAFFOLDACTIVITYTEXTREQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Activity Text Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagActivityText1Required = result == BoolResult.True ? true : false;
                            }
                            break;


                        case "SCAFFOLDLIST1LABEL":
                            data = CellData("SCAFFOLDLIST1LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList1Label = data.Item1;
                                jobsite.UseScaffoldTagList1 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST1REQUIRED":
                            data = CellData("SCAFFOLDLIST1REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List1 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList1Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDLIST2LABEL":
                            data = CellData("SCAFFOLDLIST2LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList2Label = data.Item1;
                                jobsite.UseScaffoldTagList2 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST2REQUIRED":
                            data = CellData("SCAFFOLDLIST2REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List2 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList2Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDLIST3LABEL":
                            data = CellData("SCAFFOLDLIST3LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList3Label = data.Item1;
                                jobsite.UseScaffoldTagList3 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST3REQUIRED":
                            data = CellData("SCAFFOLDLIST3REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List3 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList3Required = result == BoolResult.True ? true : false;
                            }
                            break;


                        case "SCAFFOLDLIST4LABEL":
                            data = CellData("SCAFFOLDLIST4LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList4Label = data.Item1;
                                jobsite.UseScaffoldTagList4 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST4REQUIRED":
                            data = CellData("SCAFFOLDLIST4REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List4 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList4Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDLIST5LABEL":
                            data = CellData("SCAFFOLDLIST5LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList5Label = data.Item1;
                                jobsite.UseScaffoldTagList5 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST5REQUIRED":
                            data = CellData("SCAFFOLDLIST5REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List5 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList5Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDLIST6LABEL":
                            data = CellData("SCAFFOLDLIST6LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList6Label = data.Item1;
                                jobsite.UseScaffoldTagList6 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST6REQUIRED":
                            data = CellData("SCAFFOLDLIST6REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List6 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList6Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDLIST7LABEL":
                            data = CellData("SCAFFOLDLIST7LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagList7Label = data.Item1;
                                jobsite.UseScaffoldTagList7 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDLIST7REQUIRED":
                            data = CellData("SCAFFOLDLIST7REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold List7 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagList7Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDDATE1LABEL":
                            data = CellData("SCAFFOLDDATE1LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagDate1Label = data.Item1;
                                jobsite.UseScaffoldTagDate1 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDDATE1REQUIRED":
                            data = CellData("SCAFFOLDDATE1REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Date1 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagDate1Required = result == BoolResult.True ? true : false;
                            }
                            break;


                        case "SCAFFOLDDATE2LABEL":
                            data = CellData("SCAFFOLDDATE2LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagDate2Label = data.Item1;
                                jobsite.UseScaffoldTagDate2 = data.Item1 != "";
                            }
                            break;
                        case "SCAFFOLDDATE2REQUIRED":
                            data = CellData("SCAFFOLDDATE2REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Date2 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagDate2Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDTEXT1LABEL":
                            data = CellData("SCAFFOLDTEXT1LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagText1Label = data.Item1;
                                jobsite.UseScaffoldTagText1 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDTEXT1REQUIRED":
                            data = CellData("SCAFFOLDTEXT1REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Text1 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagText1Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDTEXT2LABEL":
                            data = CellData("SCAFFOLDTEXT2LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagText2Label = data.Item1;
                                jobsite.UseScaffoldTagText2 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDTEXT2REQUIRED":
                            data = CellData("SCAFFOLDTEXT2REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Text2 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagText2Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDTEXT3LABEL":
                            data = CellData("SCAFFOLDTEXT3LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagText3Label = data.Item1;
                                jobsite.UseScaffoldTagText3 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDTEXT3REQUIRED":
                            data = CellData("SCAFFOLDTEXT3REQUIRED", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Text3 Required must be TRUE, FALSE, YES, NO or left blank.  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                jobsite.ScaffoldTagText3Required = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "SCAFFOLDYES/NO1LABEL":
                            data = CellData("SCAFFOLDYES/NO1LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagYesNo1Label = data.Item1;
                                jobsite.UseScaffoldTagYesNo1 = data.Item1 != "";
                            }
                            break;

                        case "SCAFFOLDYES/NO2LABEL":
                            data = CellData("SCAFFOLDYES/NO2LABEL", ws, row, importDict);
                            if (data != null)
                            {
                                jobsite.ScaffoldTagYesNo2Label = data.Item1;
                                jobsite.UseScaffoldTagYesNo2 = data.Item1 != "";
                            }
                            break;

                        case "BUSINESSSTREET1":
                            data = CellData("BUSINESSSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                businessAddress.Street = data.Item1;
                            }
                            break;

                        case "BUSINESSSTREET2":
                            data = CellData("BUSINESSSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                businessAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BUSINESSCITY":
                            data = CellData("BUSINESSCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                businessAddress.City = data.Item1;
                            }
                            break;

                        case "BUSINESSSTATE":
                            data = CellData("BUSINESSSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            data = CellData("BUSINESSCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    businessAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "BUSINESSZIP/POSTALCODE":
                            data = CellData("BUSINESSZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                businessAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            data = CellData("BUSINESSCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                businessAddress.Country = data.Item1;
                            }
                            break;
                        case "BUSINESSLATITUDE":
                            data = CellData("BUSINESSLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }
                                try
                                {
                                    businessAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BUSINESSLONGITUDE":
                            data = CellData("BUSINESSLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            data = CellData("BUSINESSELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (businessAddress == null)
                                {
                                    businessAddress = GetLocationAddress(jobsite, AddressTypes.Business);
                                }

                                try
                                {
                                    businessAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Business Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGSTREET1":
                            data = CellData("BILLINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }

                                billingAddress.Street = data.Item1;
                            }
                            break;

                        case "BILLINGSTREET2":
                            data = CellData("BILLINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }
                                billingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "BILLINGCITY":
                            data = CellData("BILLINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }
                                billingAddress.City = data.Item1;
                            }
                            break;

                        case "BILLINGSTATE":
                            data = CellData("BILLINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            data = CellData("BILLINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    billingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            data = CellData("BILLINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }
                                billingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            data = CellData("BILLINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }
                                billingAddress.Country = data.Item1;
                            }
                            break;
                        case "BILLINGLATITUDE":
                            data = CellData("BILLINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }
                                try
                                {
                                    billingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "BILLINGLONGITUDE":
                            data = CellData("BILLINGLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            data = CellData("BILLINGELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (billingAddress == null)
                                {
                                    billingAddress = GetLocationAddress(jobsite, AddressTypes.Billing);
                                }

                                try
                                {
                                    billingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Billing Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGSTREET1":
                            data = CellData("SHIPPINGSTREET1", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }

                                shippingAddress.Street = data.Item1;
                            }
                            break;

                        case "SHIPPINGSTREET2":
                            data = CellData("SHIPPINGSTREET2", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }
                                shippingAddress.Street1 = data.Item1;
                            }
                            break;

                        case "SHIPPINGCITY":
                            data = CellData("SHIPPINGCITY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }
                                shippingAddress.City = data.Item1;
                            }
                            break;

                        case "SHIPPINGSTATE":
                            data = CellData("SHIPPINGSTATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }

                                state = State.GetState(data.Item1);
                                if (state.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping State is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.StateID = state.StateID;
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            data = CellData("SHIPPINGCOUNTY/PARISH", ws, row, importDict);

                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }

                                county = County.GetCounty(data.Item1);
                                if (county.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping County/Parish is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    shippingAddress.CountyID = county.CountyID;
                                }
                            }
                            break;
                        case "SHIPPINGZIP/POSTALCODE":
                            data = CellData("SHIPPINGZIP/POSTALCODE", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }
                                shippingAddress.PostalCode = data.Item1;
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            data = CellData("SHIPPINGCOUNTRY", ws, row, importDict);
                            if (data != null)
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }
                                shippingAddress.Country = data.Item1;
                            }
                            break;
                        case "SHIPPINGLATITUDE":
                            data = CellData("SHIPPINGLATITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }
                                try
                                {
                                    shippingAddress.Latitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Latitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "SHIPPINGLONGITUDE":
                            data = CellData("SHIPPINGLONGITUDE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Longitude = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Longitude is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "SHIPPINGELEVATION":
                            data = CellData("SHIPPINGELEVATION", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (shippingAddress == null)
                                {
                                    shippingAddress = GetLocationAddress(jobsite, AddressTypes.Shipping);
                                }

                                try
                                {
                                    shippingAddress.Elevation = double.Parse(data.Item1);
                                }
                                catch
                                {
                                    WriteCell(row, resultCol, ws, "The value for Shipping Elevation is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "JOBTEXT1":
                            data = CellData("JOBTEXT1", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite.JobText1 = data.Item1;
                            }
                            break;

                        case "JOBTEXT2":
                            data = CellData("JOBTEXT2", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite.JobText2 = data.Item1;
                            }
                            break;

                        case "JOBYESNO1":
                            data = CellData("JOBYESNO1", ws, row, importDict);
                            result = CellBool(data, true, true);
                            switch (result)
                            {
                                case BoolResult.Invalid:
                                    WriteCell(row, resultCol, ws, "The value for Job YesNo1 is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                case BoolResult.Empty:
                                    break;
                                default:
                                    jobsite.JobYesNo1 = result == BoolResult.True;
                                    break;
                            }
                            break;

                        case "JOBYESNO2":
                            data = CellData("JOBYESNO2", ws, row, importDict);
                            result = CellBool(data, true, true);
                            switch (result)
                            {
                                case BoolResult.Invalid:
                                    WriteCell(row, resultCol, ws, "The value for Job YesNo2 is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                case BoolResult.Empty:
                                    break;
                                default:
                                    jobsite.JobYesNo2 = result == BoolResult.True;
                                    break;
                            }
                            break;

                        case "JOBLIST1":
                            data = CellData("JOBLIST1", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                JobList1 list1 = JobList1.GetJobList1(data.Item1);
                                if (list1.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    jobsite.JobList1ID = list1.JobList1ID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "The value for Job List1 is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "JOBLIST2":
                            data = CellData("JOBLIST2", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                JobList2 list2 = JobList2.GetJobList2(data.Item1);
                                if (list2.Name.ToUpper() == data.Item1.ToUpper())
                                {
                                    jobsite.JobList2ID = list2.JobList2ID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "The value for Job List2 is not valid.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "BILLINGMETHOD":
                            data = CellData("BILLINGMETHOD", ws, row, importDict);
                            string value = data.Item1.ToUpper();
                            if (jobsite.IsBillable)
                            {
                                if (ValidBillingMethods.Contains(value))
                                {
                                    switch (value)
                                    {
                                        case "ARREARS":
                                            jobsite.BillingMethod = BillingMethodType.Arrears;
                                            break;
                                        case "FATA":
                                            jobsite.BillingMethod = BillingMethodType.FirstAdvancedThenArrears;
                                            break;
                                        case "ADVANCE":
                                            jobsite.BillingMethod = BillingMethodType.Advanced;
                                            break;
                                    }
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Billing Method must be Arrears, FATA or Advance for billable jobsites.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            else if (data != null && data.Item1 != "")
                            {
                                WriteCell(row, resultCol, ws, "Billing Method is valid only for billable job sites.  ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;
                        case "FIRSTINVOICEDATE":
                            data = CellData("FIRSTINVOICEDATE", ws, row, importDict);
                            value = data.Item1.ToUpper();

                            if (jobsite.BillingMethod == BillingMethodType.Arrears)
                            {
                                if (value == "")
                                {
                                    WriteCell(row, resultCol, ws, "First Invoice Date is required for Arrears Billing.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    bool isDate = DateTime.TryParse(value, out DateTime tempDate);
                                    if (!isDate)
                                    {
                                        WriteCell(row, resultCol, ws, "First Invoice Date is not a valid date.  ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        jobsite.FirstInvoiceDate = tempDate.ToShortDateString();
                                    }
                                }
                            }
                            else
                            {
                                if (value != "")
                                {
                                    WriteCell(row, resultCol, ws, "First Invoice Date is valid only for the Arrears Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "FATAMINIMUMDAYSRENT":
                            data = CellData("FATAMINIMUMDAYSRENT", ws, row, importDict);
                            value = data.Item1.ToUpper();

                            if (jobsite.BillingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                if (value == "")
                                {
                                    WriteCell(row, resultCol, ws, "Fata Minimum Days Rent is required for FATA Billing.  ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    bool isInt = int.TryParse(value, out int tempInt);
                                    if (!isInt || tempInt < 1)
                                    {
                                        WriteCell(row, resultCol, ws, "Fata Minimum Days Rent must be an interger greater than zero for FATA billing.  ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        jobsite.MinimumDaysFata = tempInt;
                                    }
                                }
                            }
                            else
                            {
                                if (value != "")
                                {
                                    WriteCell(row, resultCol, ws, "Fata Minimum Days Rent applies only to the FATA Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "FATAALLOWOVERRIDE":
                            data = CellData("FATAALLOWOVERRIDE", ws, row, importDict);
                            result = CellBool(data, false, true);

                            if (jobsite.BillingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                if (result != BoolResult.Invalid)
                                {
                                    jobsite.AllowMinimumDayOverride = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Fata Allow Override applies only to the FATA Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "FATAINCLUDEADDITIONALCHARGES":
                            data = CellData("FATAINCLUDEADDITIONALCHARGES", ws, row, importDict);
                            result = CellBool(data, false, true);

                            if (jobsite.BillingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                if (result != BoolResult.Invalid)
                                {
                                    jobsite.IncludeUnitPricesOnMinimum = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Fata Include Additional Charges applies only to the FATA Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "ADVANCEISSUECREDITS":
                            data = CellData("ADVANCEISSUECREDITS", ws, row, importDict);
                            result = CellBool(data, false, true);

                            if (jobsite.BillingMethod == BillingMethodType.Advanced)
                            {
                                if (result != BoolResult.Invalid)
                                {
                                    jobsite.IssueCreditsForEarlyReturns = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Advance Issue Credits applies only to the Advance Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "ADVANCEMINIMUMDAYSRENT":
                            data = CellData("ADVANCEMINIMUMDAYSRENT", ws, row, importDict);
                            if (data == null || data.Item1 == "")
                            {
                                jobsite.MinimumDaysAdvanced = 0;
                            }
                            else if (jobsite.BillingMethod != BillingMethodType.Advanced)
                            {
                                WriteCell(row, resultCol, ws, "Advance Minimum Days Rent applies only to the Advance Billing Method.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else
                            {
                                bool isInt = int.TryParse(data.Item1, out int tempInt);
                                if (!isInt)
                                {
                                    WriteCell(row, resultCol, ws, "Advance Minimum Days Rent is not an integer.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.MinimumDaysAdvanced = tempInt;
                                }
                            }
                            break;
                        case "ADVANCEINVOICECYCLEDAYS":
                            data = CellData("ADVANCEINVOICECYCLEDAYS", ws, row, importDict);
                            if (data == null || data.Item1 == "")
                            {
                                if (jobsite.BillingMethod == BillingMethodType.Advanced)
                                {
                                    WriteCell(row, resultCol, ws, "Advance Invoice Cycle Days must be at least 1 for Advance billing.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            else
                            {
                                if (jobsite.BillingMethod != BillingMethodType.Advanced)
                                {

                                    WriteCell(row, resultCol, ws, "Advance Invoice Cycle Days is valid only for Advance billing.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    bool isInt = int.TryParse(data.Item1, out int tempInt);
                                    if (!isInt || tempInt <= 0)
                                    {
                                        WriteCell(row, resultCol, ws, "Advance Invoice Cycle Days must be an integer greater than zero.  ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    jobsite.CycleDaysAdvanced = tempInt;
                                }
                            }
                            break;
                        case "ADVANCECYCLECOMMONDATE":
                            data = CellData("ADVANCECYCLECOMMONDATE", ws, row, importDict);
                            result = CellBool(data, false, true);

                            if (jobsite.BillingMethod == BillingMethodType.Advanced)
                            {
                                if (result != BoolResult.Invalid)
                                {
                                    jobsite.AdvanceIsCycled = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Advance Cycle Common Date applies only to the Advance Billing Method.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "ADVANCECYCLESTARTDATE":
                            data = CellData("ADVANCECYCLESTARTDATE", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (!jobsite.AdvanceIsCycled)
                                {
                                    WriteCell(row, resultCol, ws, "Advance Cycle Start Date applies only if Advance Cycle Common Date is True or Yes.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    bool isDate = DateTime.TryParse(data.Item1, out DateTime tempDate);
                                    if (!isDate)
                                    {
                                        WriteCell(row, resultCol, ws, "Advance Cycle Start Date is not a valid date.  ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        jobsite.AwcFirstInvoiceDate = tempDate.ToShortDateString();
                                    }
                                }
                            }
                            else if (jobsite.AdvanceIsCycled)
                            {
                                WriteCell(row, resultCol, ws, "Advance Cycle Start Date is required if Advance Cycle Common Date is True or Yes.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;
                        case "DAILY/MONTHLY":
                            data = CellData("DAILY/MONTHLY", ws, row, importDict);
                            if (jobsite.IsBillable && (jobsite.BillingMethod == BillingMethodType.FirstAdvancedThenArrears || jobsite.BillingMethod == BillingMethodType.Arrears))
                            {
                                if (data == null || data.Item1 == "")
                                {
                                    WriteCell(row, resultCol, ws, "Daily / Monthly must be \"Daily\" or \"Monthly\" for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    switch (data.Item1.ToUpper())
                                    {
                                        case "DAILY":
                                            jobsite.ArrearsBillingCycle = ArrearsBillingCycleType.Daily;
                                            break;
                                        case "MONTHLY":
                                            jobsite.ArrearsBillingCycle = ArrearsBillingCycleType.Monthly;
                                            break;
                                        default:
                                            WriteCell(row, resultCol, ws, "Daily / Monthly must be \"Daily\" or \"Monthly\" for billable job sites. ",
                                                useColor: true, append: true, color: warningColor);
                                            importFailure = true;
                                            breakLoop = true;
                                            break;
                                    }
                                }
                            }
                            else if (data != null && data.Item1 != "")
                            {
                                WriteCell(row, resultCol, ws, "Daily / Monthly is valid only for billable Arrears or FATA jobsites. ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;
                        case "INVOICEEVERYNUMBEROFDAYS":
                            data = CellData("INVOICEEVERYNUMBEROFDAYS", ws, row, importDict);
                            if (data == null || data.Item1 == "")
                            {
                                if ((jobsite.BillingMethod == BillingMethodType.Arrears ||
                                    jobsite.BillingMethod == BillingMethodType.FirstAdvancedThenArrears) &&
                                    jobsite.ArrearsBillingCycle == ArrearsBillingCycleType.Daily)
                                {
                                    WriteCell(row, resultCol, ws, "Invoice Every Number of Days is required for Daily billing.  ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            else
                            {
                                if (jobsite.ArrearsBillingCycle != ArrearsBillingCycleType.Daily)
                                {
                                    WriteCell(row, resultCol, ws, "Invoice Every Number of Days is valid only for Daily billing. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    bool isInt = int.TryParse(data.Item1, out int tempInt);
                                    if (!isInt || tempInt <= 0)
                                    {
                                        WriteCell(row, resultCol, ws, "Invoice Every Number of Days must be an integer greater than zero.  ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    jobsite.CycleDaysArrears = tempInt;
                                }
                            }
                            break;
                        case "INVOICEMONTHLYCYCLETYPE":
                            data = CellData("INVOICEMONTHLYCYCLETYPE", ws, row, importDict);
                            if ((data == null || data.Item1 == "") &&
                                jobsite.ArrearsBillingCycle == ArrearsBillingCycleType.Monthly)
                            {
                                WriteCell(row, resultCol, ws, "Invoice Monthly Cycle Type is required for Monthly billing.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else if ((data != null && data.Item1 != "") && jobsite.ArrearsBillingCycle != ArrearsBillingCycleType.Monthly)
                            {
                                WriteCell(row, resultCol, ws, "Invoice Monthly Cycle Type is valid only for Monthly billing.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else if ((data != null && data.Item1 != ""))
                            {
                                switch (data.Item1.ToUpper())
                                {
                                    case "FIRST DAY":
                                        jobsite.MonthlyBillingCycle = MonthlyBillingCycleType.FirstDay;
                                        break;
                                    case "LAST DAY":
                                        jobsite.MonthlyBillingCycle = MonthlyBillingCycleType.LastDay;
                                        break;
                                    case "SPECIFIC DAY":
                                        jobsite.MonthlyBillingCycle = MonthlyBillingCycleType.SpecificDay;
                                        break;
                                    case "LAST SUNDAY":
                                        jobsite.MonthlyBillingCycle = MonthlyBillingCycleType.LastSunday;
                                        break;
                                    default:
                                        WriteCell(row, resultCol, ws, "Invoice Monthly Cycle Type must be First Day, Last Day, Last Sunday or Specific Day.  ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                }
                            }
                            break;
                        case "INVOICEMONTHLYDAYOFMONTH":
                            data = CellData("INVOICEMONTHLYDAYOFMONTH", ws, row, importDict);
                            if ((data != null && data.Item1 != "") &&
                                jobsite.MonthlyBillingCycle != MonthlyBillingCycleType.SpecificDay)
                            {
                                WriteCell(row, resultCol, ws, "Invoice Monthly Day of Month is valid only for Specific Day billing.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else if ((data == null || data.Item1 == "") &&
                                    jobsite.MonthlyBillingCycle == MonthlyBillingCycleType.SpecificDay)
                            {
                                WriteCell(row, resultCol, ws, "Invoice Monthly Day of Month is required for Specific Day billing.  ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else if (data != null && data.Item1 != "")
                            {
                                bool isInt = int.TryParse(data.Item1, out int tempInt);
                                if (!isInt || tempInt < 1 || tempInt > 31)
                                {
                                    WriteCell(row, resultCol, ws, "Invoice Monthly Day of Month must be an integer between 1 and 31. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.DayOfMonth = tempInt;
                                }
                            }
                            break;
                        case "RATEPROFILE":
                            data = CellData("RATEPROFILE", ws, row, importDict);
                            if ((data == null || data.Item1 == "") && jobsite.IsBillable)
                            {
                                WriteCell(row, resultCol, ws, "Rate Profile is required for billable job sites. ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            else if ((data != null && data.Item1 != "") && !jobsite.IsBillable)
                            {
                                WriteCell(row, resultCol, ws, "Rate Profile is only valid for billable job sites. ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;

                            }
                            else if (data != null && data.Item1 != "")
                            {
                                string tempName = data.Item1.ToUpper();
                                RateProfile profile = RateProfile.GetRateProfile(tempName, false);
                                if (profile.Name.ToUpper() != tempName)
                                {
                                    WriteCell(row, resultCol, ws, "Rate Profile not found in database. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.DefaultRateProfileID = profile.RateProfileID;
                                }
                            }
                            break;
                        case "PURCHASEORDER":
                            data = CellData("PURCHASEORDER", ws, row, importDict);
                            if (data == null || data.Item1 == "")
                            {
                                if (jobsite.IsBillable && jobsite.TrackScaffoldTags)
                                {
                                    WriteCell(row, resultCol, ws, "Purchase Order is required for billable scaffold tag tracking job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                            }
                            else
                            {
                                order = Order.GetOrder(data.Item1);
                                if (order.Number.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Purchase Order not found in database. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                            }
                            break;
                        case "WORKORDER":
                            data = CellData("WORKORDER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                if (order == null)
                                {
                                    WriteCell(row, resultCol, ws, "Work Order is not valid without a Purchase Order. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                                else
                                {
                                    Order workorder = Order.GetWorkOrder(order.Number, data.Item1);
                                    if (workorder.Number.ToUpper() != data.Item1.ToUpper())
                                    {
                                        WriteCell(row, resultCol, ws, "[Purchase Order, Work Order] not found in the database. ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                    else if (!order.IsActive || !workorder.IsActive)
                                    {
                                        WriteCell(row, resultCol, ws, "Purchase Order & Work Order must both be active. ",
                                            useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                        break;
                                    }
                                    else
                                    {
                                        jobsite.OrderID = workorder.OrderID;
                                    }
                                }
                            }
                            else if (order != null && order.IsActive)
                            {
                                jobsite.OrderID = order.OrderID;
                            }
                            else if (order != null)
                            {
                                WriteCell(row, resultCol, ws, "Purchase Order must both be active. ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                                break;
                            }
                            break;
                        case "DEFAULTNEWSHIPMENTSBILLABLE":
                            data = CellData("DEFAULTNEWSHIPMENTSBILLABLE", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, false, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Default New Shipments Billable must be True, False, Yes, or No for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.DefaultShipmentToBillable = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Default New Shipments Billable is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "HIDEZEROCHARGES":
                            data = CellData("HIDEZEROCHARGES", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, true, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Default New Shipments Billable must be True, False, Yes, No or blank for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.HideZeroInvoiceItems = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Hide Zero Charges is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXRATE1":
                            data = CellData("TAXRATE1", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string tempName = data.Item1.ToUpper();
                                TaxRate taxRate = TaxRate.GetTaxRate(tempName);
                                if (taxRate.Name.ToUpper() == tempName)
                                {
                                    jobsite.JobTax1ID = taxRate.TaxRateID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Tax Rate1 not found in the database. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXRATE2":
                            data = CellData("TAXRATE2", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                string tempName = data.Item1.ToUpper();
                                TaxRate taxRate = TaxRate.GetTaxRate(tempName);
                                if (taxRate.Name.ToUpper() == tempName)
                                {
                                    jobsite.JobTax2ID = taxRate.TaxRateID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Tax Rate2 not found in the database. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXCONSUMABLESANDPRODUCTSALES":
                            data = CellData("TAXCONSUMABLESANDPRODUCTSALES", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, true, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Tax Consumables and Product Sales must be True, False, Yes, No or blank for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.ConsumablesTaxable = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Tax Consumables and Product Sales is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXRENT":
                            data = CellData("TAXRENT", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, true, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Tax Rent must be True, False, Yes, No or blank for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.RentIsTaxable = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Tax Rent is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXSERVICETICKETDAMAGE":
                            data = CellData("TAXSERVICETICKETDAMAGE", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, true, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Tax Service Ticket Damage must be True, False, Yes, No or blank for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.ServiceTicketDamageChargeTaxable = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Tax Service Ticket Damage is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "TAXDELIVERY&RETURN":
                            data = CellData("TAXDELIVERY&RETURN", ws, row, importDict);
                            if (jobsite.IsBillable)
                            {
                                result = CellBool(data, true, true);

                                if (result == BoolResult.Invalid)
                                {
                                    WriteCell(row, resultCol, ws, "Tax Delivery & Return must be True, False, Yes, No or blank for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite.DeliveryChargesTaxable = result == BoolResult.True;
                                }
                            }
                            else
                            {
                                if (data != null && data.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Tax Delivery & Return is valid only for billable job sites. ",
                                        useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        default:
                            Debug.WriteLine("Unexpected case = {0} ", NormStr(configCol.ColumnHeading));
                            break;
                    }
                }
            }
            if (!importFailure)
            {
                object importObject = null;
                if (jobsite.IsSavable)
                {
                    importObject = jobsite;
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject, parentName, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in jobsite.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportScaffoldJobList(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            StockingLocation jobsite = null;
            string listName = "";
            string value = "";
            string description = "";
            int listColumn = config.ImportColumns.Where(x => NormStr(x.ColumnHeading) == "LIST").First().Column;

            Tuple<string, ImportConfig.ImportColunnType, int> activityData;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            activityData = CellData("JOBSITENAME", ws, row, importDict);

                            if (activityData != null)
                            {
                                string jobName = activityData.Item1;
                                if (jobName != "")
                                {
                                    jobsite = StockingLocation.GetStockingLocation(jobName, false);
                                }
                                if (jobName == "" || jobsite.Name.ToUpper() != jobName.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Jobsite name is not valid. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "LIST":
                            activityData = CellData("LIST", ws, row, importDict);
                            if (activityData != null)
                            {
                                string name = NormStr(activityData.Item1);
                                if (ValidScaffoldJobListNames.Contains(name))
                                {
                                    listName = name;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, string.Format("List value {0} is not recognized. Unable to import his value. ", listName), useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "VALUE":
                            activityData = CellData("VALUE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                value = activityData.Item1;
                                if (value == "")
                                {
                                    importFailure = true;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, string.Format("The Value is required; it can't be left blank. ", listName), useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;
                        case "DESCRIPTION":
                            activityData = CellData("DESCRIPTION", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                description = activityData.Item1;
                            }
                            else
                            {
                                description = "";
                            }
                            break;


                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                object importObject1 = null;
                object importObject2 = null;
                ListNumberTypes type = ListNumberTypes.None;

                if (listName == "SCAFFOLDTAGACTIVITYTYPE")
                {
                    var activityType = ScaffoldTagActivityType.NewScaffoldTagActivityType(jobsite.StockingLocationID);
                    activityType.Name = value;
                    activityType.Description = description;
                    if (!jobsite.UseScaffoldTagActivities)
                    {
                        WriteCell(row, resultCol, ws, "Activities are not enabled for job site. ", useColor: true, append: true, color: warningColor);
                        importFailure = true;
                    }
                    else if (activityType.IsSavable)
                    {
                        importObject1 = activityType;

                    }
                    else
                    {
                        foreach (var brokenRule in activityType.BrokenRulesCollection.Take(maxBrokenRulesReported))
                        {
                            WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                        }
                        importFailure = true;
                    }
                }
                else if (listName.Left(20) == "SCAFFOLDACTIVITYLIST")
                {
                    if (!jobsite.UseScaffoldTagActivities)
                    {
                        WriteCell(row, resultCol, ws, "Activities are not enabled for job site. ", useColor: true, append: true, color: warningColor);
                        importFailure = true;
                    }
                    else
                    {
                        switch (listName)
                        {
                            case "SCAFFOLDACTIVITYLIST1":
                                if (jobsite.UseScaffoldTagActivityList1)
                                {
                                    type = ListNumberTypes.ActivityList1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Scaffold Activity List 1 is not enabled for job site. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                break;

                            case "SCAFFOLDACTIVITYLIST2":
                                if (jobsite.UseScaffoldTagActivityList2)
                                {
                                    type = ListNumberTypes.ActivityList2;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Scaffold Activity List 2 is not enabled for job site. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                break;

                            case "SCAFFOLDACTIVITYLIST3":
                                if (jobsite.UseScaffoldTagActivityList3)
                                {
                                    type = ListNumberTypes.ActivityList3;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Scaffold Activity List 3 is not enabled for job site. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                break;

                            default:
                                WriteCell(row, resultCol, ws, "Scaffold Activity List number not recognized. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                break;
                        }
                    }

                    if (!importFailure)
                    {
                        var listItem = ScaffoldTagJobListItem.NewScaffoldTagJobListItem(jobsite.StockingLocationID, type);
                        listItem.Name = value;
                        listItem.Description = description;
                        if (listItem.IsSavable)
                        {
                            importObject2 = listItem;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                    }
                }
                else if (listName.Left(12) == "SCAFFOLDLIST")
                {
                    switch (listName)
                    {
                        case "SCAFFOLDLIST1":
                            type = ListNumberTypes.JobList1;
                            break;
                        case "SCAFFOLDLIST2":
                            type = ListNumberTypes.JobList2;
                            break;
                        case "SCAFFOLDLIST3":
                            type = ListNumberTypes.JobList3;
                            break;
                        case "SCAFFOLDLIST4":
                            type = ListNumberTypes.JobList4;
                            break;
                        case "SCAFFOLDLIST5":
                            type = ListNumberTypes.JobList5;
                            break;
                        case "SCAFFOLDLIST6":
                            type = ListNumberTypes.JobList6;
                            break;
                        case "SCAFFOLDLIST7":
                            type = ListNumberTypes.JobList7;
                            break;

                        default:
                            importFailure = true;
                            break;
                    }

                    if (!importFailure)
                    {
                        var listItem = ScaffoldTagJobListItem.NewScaffoldTagJobListItem(jobsite.StockingLocationID, type);
                        listItem.Name = value;
                        listItem.Description = description;
                        if (listItem.IsSavable)
                        {
                            importObject2 = listItem;
                        }
                        else
                        {
                            foreach (var brokenRule in listItem.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                            }
                            importFailure = true;
                        }
                    }

                }

                if (!importFailure)
                {
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>(importObject1, importObject2, null, row, resultCol, null);
                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
            }
            return importFailure;
        }

        private bool ImportRequest(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, AvontusUser user,
            Dictionary<string, List<string>> userDic,
            List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            Request request = null;

            StockingLocation jobsite = null;
            GlobalRequestListItem listItem = null;
            bool saveJobsite = false;

            Tuple<string, ImportConfig.ImportColunnType, int> data;
            bool breakLoop = false;
            BoolResult result = BoolResult.Invalid;
            importFailure = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            data = CellData("JOBSITENAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                jobsite = StockingLocation.GetStockingLocation(data.Item1, false);
                                if (jobsite.Name.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Job Site not found ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else if (!jobsite.TrackRequests)
                                {
                                    WriteCell(row, resultCol, ws, "Job Site doesn't track Requests ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    var configuredUser = StockingLocationUserAssignment.GetStockingLocationUserAssignment(user.UserID, jobsite.StockingLocationID);
                                    if (configuredUser.UserName != user.Username)
                                    {
                                        WriteCell(row, resultCol, ws, "Running User must be a valid Assign To User for Jobsite. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        RequestStatusCollection statuses = RequestStatusCollection.GetRequestStatusCollection(jobsite.StockingLocationID, ActiveStatus.Both);
                                        if (statuses.Count == 0)
                                        {

                                            statuses.CreateDefaultStatuses(jobsite.StockingLocationID);
                                            saveJobsite = true;
                                        }

                                        if (saveJobsite)
                                        {
                                            try
                                            {
                                                jobsite.Save();
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Unable to change Job Site to track requests and to initialize default statuses ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                                breakLoop = true;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Job Site Name not provided ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;


                        case "REQUESTNUMBER":
                            data = CellData("REQUESTNUMBER", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                request = Request.NewRequest(false);
                                request.RequestNumber = data.Item1;
                                request.ParentJobSiteID = jobsite.StockingLocationID;
                                request.CreatedByUserID = user.UserID;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Request Number not provided ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "NOTES":
                            data = CellData("NOTES", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                request.Notes = data.Item1;
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Notes are not provided ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "ORDER":
                            // Order here is a text field; it's not a reference to an Order in 
                            // the Orders table
                            data = CellData("ORDER", ws, row, importDict);
                            request.OrderNumber = data.Item1 ?? "";
                            break;

                        case "STATUS":
                            data = CellData("STATUS", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {

                                string revStatus = data.Item1.ToUpper();
                                switch (revStatus)
                                {
                                    case "INITIAL REQUEST":
                                        revStatus = "1. Initial Request";
                                        break;
                                    case "IN PROGRESS":
                                        revStatus = "2. In Progress";
                                        break;

                                    case "CLOSED":
                                        revStatus = "3. Closed";
                                        break;
                                    default:
                                        revStatus = data.Item1;
                                        break;

                                }

                                RequestStatus stat = RequestStatus.GetRequestStatus(revStatus, jobsite.StockingLocationID);
                                if (stat.Name.ToUpper() == revStatus.ToUpper())
                                {
                                    request.StatusID = stat.RequestStatusID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Status is not recognized ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;

                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Status is not provided ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;

                            }
                            break;

                        case "ASSIGNEDTOUSERNAME":
                            data = CellData("ASSIGNEDTOUSERNAME", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                AvontusUser assignedUser = AvontusUser.GetUser(data.Item1);
                                if (assignedUser.Username.ToUpper() != data.Item1.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Assigned To Username is not recognized ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    var jobUser = StockingLocationUserAssignment.GetStockingLocationUserAssignment(assignedUser.UserID, jobsite.StockingLocationID);
                                    if (jobUser.UserName != assignedUser.Username)
                                    {
                                        WriteCell(row, resultCol, ws, "Assigned To User is not configured for assignment in the job site. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    request.AssignedToUserID = assignedUser.UserID;
                                }
                            }
                            break;

                        case "REQUESTDATE":
                            data = CellData("REQUESTDATE", ws, row, importDict);
                            try
                            {
                                request.RequestDate = SmartDate.Parse(data.Item1);
                            }
                            catch
                            {
                                WriteCell(row, resultCol, ws, "Request Date is invalid ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "DATENEEDED":
                            data = CellData("DATENEEDED", ws, row, importDict);
                            try
                            {
                                request.DateNeeded = SmartDate.Parse(data.Item1);
                            }
                            catch
                            {
                                WriteCell(row, resultCol, ws, "DateNeeded is invalid ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            break;

                        case "REQUESTLIST1":
                            data = CellData("REQUESTLIST1", ws, row, importDict);
                            listItem = null;
                            if (data != null && data.Item1 != "")
                            {
                                var list1 = GlobalRequestList.GetGlobalRequestList(Guid.Empty, ActiveStatus.Both, false, GlobalRequestListNumberTypes.GlobalList1);
                                foreach (var item in list1)
                                {
                                    if (item.Name.ToUpper() == data.Item1.ToUpper())
                                    {
                                        listItem = item;
                                        break;
                                    }
                                }

                                if (listItem != null)
                                {
                                    request.GlobalRequestList1ID = listItem.GlobalRequestListItemID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Request List 1 is not recognized ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "REQUESTLIST2":
                            data = CellData("REQUESTLIST2", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                listItem = null;

                                var list2 = GlobalRequestList.GetGlobalRequestList(Guid.Empty, ActiveStatus.Both, false, GlobalRequestListNumberTypes.GlobalList2);
                                foreach (var item in list2)
                                {
                                    if (item.Name.ToUpper() == data.Item1.ToUpper())
                                    {
                                        listItem = item;
                                        break;
                                    }
                                }

                                if (listItem != null)
                                {
                                    request.GlobalRequestList2ID = listItem.GlobalRequestListItemID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Request List 2 is not recognized ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;
                        case "REQUESTLIST3":
                            data = CellData("REQUESTLIST3", ws, row, importDict);
                            if (data != null && data.Item1 != "")
                            {
                                listItem = null;

                                var list3 = GlobalRequestList.GetGlobalRequestList(Guid.Empty, ActiveStatus.Both, false, GlobalRequestListNumberTypes.GlobalList3);
                                foreach (var item in list3)
                                {
                                    if (item.Name.ToUpper() == data.Item1.ToUpper())
                                    {
                                        listItem = item;
                                        break;
                                    }
                                }

                                if (listItem != null)
                                {
                                    request.GlobalRequestList3ID = listItem.GlobalRequestListItemID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Request List 3 is not recognized ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                            }
                            break;

                        case "REQUESTTEXT1":
                            data = CellData("REQUESTTEXT1", ws, row, importDict);
                            if (data != null & data.Item1 != "")
                            {
                                request.GlobalRequestText1 = data.Item1;
                            }
                            break;

                        case "REQUESTTEXT2":
                            data = CellData("REQUESTTEXT2", ws, row, importDict);
                            if (data != null & data.Item1 != "")
                            {
                                request.GlobalRequestText2 = data.Item1;
                            }
                            break;

                        case "REQUESTYESNO1":
                            data = CellData("REQUESTYESNO1", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Request YesNo1  must be TRUE, FALSE, YES, NO or blank .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                request.GlobalRequestYesNo1 = result == BoolResult.True ? true : false;
                            }
                            break;

                        case "REQUESTYESNO2":
                            data = CellData("REQUESTYESNO2", ws, row, importDict);
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "Request YesNo2  must be TRUE, FALSE, YES, NO or blank .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                request.GlobalRequestYesNo2 = result == BoolResult.True ? true : false;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                if (request.IsSavable)
                {
                    Tuple<object, object, object, int, int, object>
                        tuple = new Tuple<object, object, object, int, int, object>
                            (request, null, null, row, resultCol, null);

                    importedObjects.Add(tuple);
                    WriteCell(row, resultCol, ws, Validated);
                }
                else
                {
                    foreach (var brokenRule in request.BrokenRulesCollection.Take(maxBrokenRulesReported))
                    {
                        WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                    }
                    importFailure = true;
                }
            }
            return importFailure;
        }

        private bool ImportScaffoldTag(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws, ImportConfig config,
            int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            Tuple<Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, int>> importerTuple,
            Dictionary<string, Guid> savedOrders, Dictionary<string, Guid> savedWorkOrders, Dictionary<string, bool> tagsOnSheet)
        {
            bool importFailure = false;
            bool tagExists = false;

            Tuple<string, ImportConfig.ImportColunnType, int> tagData;
            StockingLocation jobsite = null;
            string orderNumber = "";
            string workOrderNumber = "";
            Order order = null;
            Guid orderID = Guid.Empty;
            Order workOrder = null;
            Guid workOrderID = Guid.Empty;
            BoolResult result = BoolResult.Invalid;
            bool orderCreated = false;
            bool workOrderCreated = false;
            string combinedOrderNumbers = "";
            string jobName = "";
            ScaffoldTag tag = null;
            ScaffoldTagStatuses saveStatus = ScaffoldTagStatuses.AllOrNone;
            int tagColumn = -1;

            foreach (var configCol in config.ImportColumns)
            {
                switch (NormStr(configCol.ColumnHeading))
                {
                    case "JOBSITENAME":
                        tagData = CellData("JOBSITENAME", ws, row, importDict);

                        jobName = tagData.Item1;

                        jobsite = StockingLocation.GetStockingLocation(jobName, false);
                        if (jobName == "" || jobsite.Name.ToUpper() != jobName.ToUpper())
                        {
                            WriteCell(row, resultCol, ws, "Jobsite name is not valid. ", useColor: true, append: true, color: warningColor);
                            importFailure = true;
                        }

                        break;
                    case "PURCHASEORDER":
                        tagData = CellData("PURCHASEORDER", ws, row, importDict);

                        orderNumber = tagData == null ? "" : tagData.Item1;

                        if (orderNumber != "")
                        {
                            if (!savedOrders.ContainsKey(orderNumber))
                            {
                                order = Order.GetOrder(orderNumber);
                                orderID = order.OrderID;
                                if (order.Number.TrimEnd() != orderNumber)
                                {
                                    order = Order.NewPurchaseOrder(jobsite.BusinessPartnerID);
                                    orderID = order.OrderID;
                                    order.Number = orderNumber;
                                    orderCreated = true;

                                }
                                else
                                {
                                    // Save the order if it was found in the DB
                                    // If it was found in the sheet (previous row) we 
                                    // save it only if the tag is saved
                                    savedOrders.Add(order.Number, order.OrderID);
                                }
                            }
                            else
                            {
                                orderID = savedOrders[orderNumber];
                            }
                        }
                        break;
                    case "WORKORDER":
                        tagData = CellData("WORKORDER", ws, row, importDict);
                        workOrderNumber = (tagData == null || tagData.Item1 == "") ? "" : tagData.Item1;
                        {
                            // Can't have a work order without a parent order
                            if (orderNumber == "" && workOrderNumber != "")
                            {
                                WriteCell(row, resultCol, ws, "Work Orders must have parent Orders. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else if (workOrderNumber != "")
                            {
                                combinedOrderNumbers = orderNumber + "\t" + workOrderNumber;
                                if (!savedWorkOrders.ContainsKey(combinedOrderNumbers))
                                {
                                    workOrder = Order.GetWorkOrder(orderNumber, workOrderNumber);
                                    workOrderID = workOrder.OrderID;
                                    if (workOrder.Number != workOrderNumber)
                                    {

                                        workOrder = Order.NewWorkOrder(jobsite.BusinessPartnerID, orderID);
                                        workOrderID = workOrder.OrderID;
                                        workOrder.Number = workOrderNumber;
                                        workOrderCreated = true;
                                    }
                                    else
                                    {

                                        // Save the work order if it was found in the DB
                                        // If it was found in the sheet (previous row) we 
                                        // save it only if the tag is saved
                                        savedWorkOrders.Add(combinedOrderNumbers, workOrder.OrderID);
                                    }
                                    workOrderID = workOrder.OrderID;
                                }
                                else
                                {
                                    workOrderID = savedWorkOrders[combinedOrderNumbers];
                                }
                            }
                        }
                        break;
                    case "STATUS":
                        tagData = CellData("STATUS", ws, row, importDict);
                        if (tagData.Item1 == "")
                        {
                            WriteCell(row, resultCol, ws, "Tag Status is required. ", useColor: true, append: true, color: warningColor);
                            importFailure = true;
                        }

                        int statusID = -1;


                        bool statusFound = importerTuple.Item4.TryGetValue(tagData.Item1.ToUpper(), out statusID);
                        if (statusFound)
                        {
                            saveStatus = (ScaffoldTagStatuses)statusID;
                        }
                        else
                        {
                            WriteCell(row, resultCol, ws, "Status not recognized ", useColor: true, append: true, color: warningColor);
                            importFailure = true;
                        }
                        break;

                    case "TAG":
                        bool problemNoted = false;
                        tagData = CellData("TAG", ws, row, importDict);
                        tagColumn = tagData.Item3 + 1;
                        if (tagData.Item1 == "")
                        {
                            WriteCell(row, resultCol, ws, "Tag identifier is not valid. ", useColor: true, append: true, color: warningColor);
                            importFailure = true;
                            problemNoted = true;
                        }
                        else
                        {
                            string tagKey = jobName + "|" + tagData.Item1;
                            if (tagsOnSheet.ContainsKey(tagKey))
                            {
                                WriteCell(row, resultCol, ws, "[Job Site Name, Tag] can not be duplicated on the import sheet. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                problemNoted = true;
                            }
                            else
                            {
                                tagsOnSheet.Add(tagKey, true);
                            }
                        }

                        Guid tempID = Guid.Empty;
                        tagExists = ScaffoldTagExistance.Exists(tagData.Item1, jobsite.Name, out tempID);

                        if (!problemNoted)
                        {
                            if (tagExists)
                            {
                                WriteCell(row, resultCol, ws, "The Jobsite already has this Tag identifier in the database. Can't import ",
                                    useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }

                            else
                            {
                                tag = ScaffoldTag.NewScaffoldTag(jobsite.StockingLocationID);
                                tag.Tag = tagData.Item1;
                                tag.ScaffoldTagStatus = saveStatus;


                                tag.OrderID = workOrderID != Guid.Empty ? workOrderID : orderID;

                                if (workOrder != null && orderID == Guid.Empty)
                                {
                                    WriteCell(row, resultCol, ws, "Scaffold Tag cannot have a Work Order without the corresponding Order. Can't import ",
                                    useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                        }
                        break;

                    default:
                        tagData = CellData(NormStr(configCol.ColumnHeading), ws, row, importDict);
                        string data = tagData.Item1.ToUpper();

                        if (tagData != null && tag != null)
                        {
                            try
                            {
                                switch (NormStr(configCol.ColumnHeading))
                                {
                                    case "PLANNEDLOADDATE":
                                        tag.PlannedLoadDate = tagData.Item1;
                                        break;
                                    case "PLANNEDBUILDDATE":
                                        tag.PlannedBuildDate = tagData.Item1;
                                        break;
                                    case "PLANNEDDISMANTLEDATE":
                                        tag.PlannedDismantleDate = tagData.Item1;
                                        break;
                                    case "ACTUALLOADDATE":
                                        tag.ActualLoadDate = tagData.Item1;
                                        break;
                                    case "ACTUALBUILDDATE":
                                        tag.ActualBuildDate = tagData.Item1;
                                        break;
                                    case "ACTUALDISMANTLEDATE":
                                        tag.ActualDismantleDate = tagData.Item1;
                                        break;
                                    case "LOCATIONNOTES":
                                        tag.LocationNotes = tagData.Item1;
                                        break;
                                    case "NOTES":
                                        tag.Notes = tagData.Item1;
                                        break;
                                    case "PRIORITY":
                                        if (tagData.Item1 != "")
                                        {
                                            Priority priority = Priority.GetPriority(tagData.Item1);
                                            if (priority.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.PriorityID = priority.PriorityID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Priority not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "PROJECT":
                                        if (tagData.Item1 != "")
                                        {
                                            Project project = Project.GetProject(tagData.Item1);
                                            if (project.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.ProjectID = project.ProjectID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Project not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "STEP":
                                        if (tagData.Item1 != "")
                                        {
                                            ScaffoldTagStep step = ScaffoldTagStep.GetScaffoldTagStep(tagData.Item1);
                                            if (step.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.ScaffoldTagStepID = step.ScaffoldTagStepID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Step not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "REQUESTEDBY":
                                        switch (data)
                                        {
                                            case "":
                                                tag.RequestedBy = RequestedByEntity.AllOrNone;
                                                break;
                                            case "CUSTOMER":
                                                tag.RequestedBy = RequestedByEntity.Customer;
                                                break;
                                            case "INTERNAL":
                                                tag.RequestedBy = RequestedByEntity.Internal;
                                                break;
                                            default:
                                                tag.RequestedBy = RequestedByEntity.AllOrNone;
                                                string msg = "Requested By must be Customer, Internal or blank. ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                                break;
                                        }
                                        break;

                                    case "REPRESENTATIVE":
                                        if (data != "")
                                        {
                                            AvontusUser repUser = AvontusUser.GetUser(data);
                                            if (repUser.Username.ToUpper() == data)
                                            {
                                                if (repUser.UserType == UserTypes.Customer &&
                                                    repUser.RelatedID == jobsite.BusinessPartnerID)
                                                {
                                                    tag.CustomerRepID = repUser.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Representative is not a valid Customer user. ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Representative is not a valid Customer user.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "REQUESTOR":
                                        if (data != "")
                                        {
                                            AvontusUser requestUser = AvontusUser.GetUser(data);
                                            if (requestUser.Username.ToUpper() == data)
                                            {
                                                if (requestUser.UserType == UserTypes.Customer &&
                                                    requestUser.RelatedID == jobsite.BusinessPartnerID)
                                                {
                                                    tag.CustomerRequestorID = requestUser.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Requestor is not assigned to the customer for the job site. ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Requestor is not a valid Customer user.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "FOREMAN":
                                        if (data != "")
                                        {
                                            AvontusUser user = AvontusUser.GetUser(data);
                                            if (user.Username.ToUpper() == data)
                                            {
                                                if (user.UserType == UserTypes.Internal)
                                                {
                                                    tag.ForemanID = user.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Forman is not a valid Quantify user.  ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Foreman not found.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "BUILDER":
                                        if (data != "")
                                        {
                                            AvontusUser user = AvontusUser.GetUser(data);
                                            if (user.Username.ToUpper() == data)
                                            {
                                                if (user.UserType == UserTypes.Internal)
                                                {
                                                    tag.BuilderID = user.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Builder is not a valid Quantify user.  ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Builder not found.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "DISMANTLER":
                                        if (data != "")
                                        {
                                            AvontusUser user = AvontusUser.GetUser(data);
                                            if (user.Username.ToUpper() == data)
                                            {
                                                if (user.UserType == UserTypes.Internal)
                                                {
                                                    tag.DismantlerID = user.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Dismantler is not a valid Quantify user.  ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Dismantler not found.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "INSPECTOR":
                                        if (data != "")
                                        {
                                            AvontusUser user = AvontusUser.GetUser(data);
                                            if (user.Username.ToUpper() == data)
                                            {
                                                if (user.UserType == UserTypes.Internal)
                                                {
                                                    tag.InspectorID = user.UserID;
                                                }
                                                else
                                                {
                                                    string msg = "Inspector is not a valid Quantify user.  ";
                                                    WriteCell(row, resultCol, ws, msg, append: true);
                                                    importFailure = true;
                                                }
                                            }
                                            else
                                            {
                                                string msg = "Inspector not found.  ";
                                                WriteCell(row, resultCol, ws, msg, append: true);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "LENGTH":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.Length = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Length is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "WIDTH":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.Width = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Width is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "HEIGHT":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.Height = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Height is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "NOOFLEGS":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.NumberOfLegs = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "No Of Legs is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "NOOFDECKS":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.NumberOfDecks = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "No Of Decks is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "BASEELEVATION":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.BaseElevation = double.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Base Elevation is not a valid floating point number ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST1":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList1);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList1ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 1 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST2":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList2);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList2ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 2 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST3":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList3);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList3ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 3 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST4":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList4);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList4ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 4 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST5":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList5);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList5ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 5 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDLIST6":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList6);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList6ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 6 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;


                                    case "SCAFFOLDLIST7":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            var listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID, tagData.Item1.ToUpper(), ListNumberTypes.JobList7);

                                            if (listItem.Name.ToUpper() == tagData.Item1.ToUpper())
                                            {
                                                tag.JobList7ID = listItem.ScaffoldTagJobListItemID;
                                            }
                                            else
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold List 7 value not recognized ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDDATE1":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.JobDate1 = SmartDate.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold Date 1 is invalid ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;
                                    case "SCAFFOLDDATE2":
                                        if (tagData != null && tagData.Item1 != "")
                                        {
                                            try
                                            {
                                                tag.JobDate2 = SmartDate.Parse(tagData.Item1);
                                            }
                                            catch
                                            {
                                                WriteCell(row, resultCol, ws, "Scaffold Date 2 is invalid ", useColor: true, append: true, color: warningColor);
                                                importFailure = true;
                                            }
                                        }
                                        break;

                                    case "SCAFFOLDTEXT1":
                                        tag.JobText1 = tagData.Item1;
                                        break;

                                    case "SCAFFOLDTEXT2":
                                        tag.JobText2 = tagData.Item1;
                                        break;

                                    case "SCAFFOLDTEXT3":
                                        tag.JobText3 = tagData.Item1;
                                        break;

                                    case "SCAFFOLDYES/NO1":
                                        result = CellBool(tagData, true, true);
                                        if (result == BoolResult.Invalid)
                                        {
                                            WriteCell(row, resultCol, ws, "Scaffold YesNo1  must be TRUE, FALSE, YES, NO or blank .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                            importFailure = true;
                                        }
                                        else
                                        {
                                            tag.JobYesNo1 = result == BoolResult.True ? true : false;
                                        }
                                        break;

                                    case "SCAFFOLDYES/NO2":
                                        result = CellBool(tagData, true, true);
                                        if (result == BoolResult.Invalid)
                                        {
                                            WriteCell(row, resultCol, ws, "Scaffold YesNo2  must be TRUE, FALSE, YES, NO or blank .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                            importFailure = true;
                                        }
                                        else
                                        {
                                            tag.JobYesNo2 = result == BoolResult.True ? true : false;
                                        }
                                        break;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message == "String value can not be converted to a date")
                                {
                                    WriteCell(row, resultCol, ws, string.Format("The {0} column contains an incorrectly formatted date ", configCol.ColumnHeading),
                                        useColor: true, append: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, string.Format("Unable to import column {0}. Reason is: {1}", configCol.ColumnHeading, ex.Message),
                                        useColor: true, append: true, color: warningColor);
                                }

                                importFailure = true;
                            }
                        }
                        break;
                }
            }
            if (importFailure)
            {
                ((Excel.Range)ws.Cells[row, tagColumn]).Interior.Color = warningColor;
            }

            if (!importFailure && tag.IsSavable)
            {
                Tuple<object, object, object, int, int, object>
                    listItem = new Tuple<object, object, object, int, int, object>(tag,
                    orderCreated ? order : null, workOrderCreated ? workOrder : null, row, resultCol, null);

                importedObjects.Add(listItem);

                if (orderCreated)
                {
                    savedOrders.Add(order.Number, order.OrderID);
                }

                if (workOrderCreated)
                {
                    savedWorkOrders.Add(combinedOrderNumbers, workOrder.OrderID);
                }
                WriteCell(row, resultCol, ws, Validated);
            }
            else if (tag != null && !tagExists)
            {
                foreach (var brokenRule in tag.BrokenRulesCollection.Take(maxBrokenRulesReported))
                {
                    WriteCell(row, resultCol, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), useColor: true, append: true, color: warningColor);
                }
                importFailure = true;
            }

            return importFailure;
        }

        private bool ImportScaffoldTagActivity(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            Tuple<Dictionary<string, Guid>, Dictionary<Tuple<Guid, string>, Guid>, Dictionary<Tuple<Guid, Int16, string>, Guid>> importerTuple)
        {
            // importerTuple contains lookups for Shift, ScaffoldTagActivityType and ScaffoldTagJobList
            bool importFailure = false;
            int summaryColumn = config.ImportColumns.Where(x => NormStr(x.ColumnHeading) == "SUMMARY").First().Column;
            ScaffoldTagActivity activity = null;
            StockingLocation jobsite = null;
            Guid tagID = Guid.Empty;

            Tuple<string, ImportConfig.ImportColunnType, int> activityData;
            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            activityData = CellData("JOBSITENAME", ws, row, importDict);

                            if (activityData != null)
                            {
                                string jobName = activityData.Item1;
                                if (jobName != "")
                                {
                                    jobsite = StockingLocation.GetStockingLocation(jobName, false);
                                }
                                if (jobName == "" || jobsite.Name.ToUpper() != jobName.ToUpper())
                                {
                                    WriteCell(row, resultCol, ws, "Jobsite name is not valid. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "TAG":
                            activityData = CellData("TAG", ws, row, importDict);
                            if (activityData != null)
                            {
                                string tagName = activityData.Item1;
                                if (jobsite == null)
                                {
                                    WriteCell(row, resultCol, ws, "Jobsite is not valid. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;
                                }
                                bool tagExists = tagName != "" ? ScaffoldTagExistance.Exists(tagName, jobsite.Name, out tagID) : false;
                                if (!tagExists)
                                {
                                    WriteCell(row, resultCol, ws, "Tag is not valid. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                    break;   // Can't do the rest for this activity
                                }
                                else
                                {
                                    activity = ScaffoldTagActivity.NewScaffoldTagActivity(tagID);
                                }
                            }

                            break;
                        case "SCAFFOLDTAGACTIVITYTYPE":
                            activityData = CellData("SCAFFOLDTAGACTIVITYTYPE", ws, row, importDict);
                            Guid typeID;

                            if (activityData == null || activityData.Item1 == "")
                            {
                                WriteCell(row, resultCol, ws, "Scaffold Tag Activity Type is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                Tuple<Guid, string> typeKey = new Tuple<Guid, string>(jobsite.StockingLocationID, activityData.Item1.ToUpper());
                                bool typeFound = importerTuple.Item2.TryGetValue(typeKey, out typeID);
                                if (typeFound)
                                {
                                    activity.ScaffoldTagActivityTypeID = typeID;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Activity Type is not valid. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "SUMMARY":
                            activityData = CellData("SUMMARY", ws, row, importDict);
                            if (activityData != null)
                            {
                                string summary = activityData.Item1;
                                if (summary == "")
                                {
                                    WriteCell(row, resultCol, ws, "Summary is required for a Scaffold Tag Activity. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                                else
                                {
                                    activity.Summary = summary;
                                }
                            }
                            break;
                        case "DETAIL":
                            activityData = CellData("DETAIL", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                activity.Detail = activityData.Item1.Left(255);
                            }
                            break;
                        case "SHIFT":
                            activityData = CellData("SHIFT", ws, row, importDict);
                            if (activityData != null)
                            {
                                Guid shiftID;
                                bool shiftFound = importerTuple.Item1.TryGetValue(activityData.Item1.ToUpper(), out shiftID);
                                if (shiftFound)
                                {
                                    activity.ShiftID = shiftID;
                                }
                                else if (activityData.Item1 != "")
                                {
                                    WriteCell(row, resultCol, ws, "Shift is not valid ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "REQUESTDATE":
                            activityData = CellData("REQUESTDATE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                DateTime tempDate;
                                bool isDate = DateTime.TryParse(activityData.Item1, out tempDate);
                                if (isDate)
                                {
                                    activity.RequestDate = activityData.Item1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Request Date is not in valid date format. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "PLANNEDDATE":
                            activityData = CellData("PLANNEDDATE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                DateTime tempDate;
                                bool isDate = DateTime.TryParse(activityData.Item1, out tempDate);
                                if (isDate)
                                {
                                    activity.PlannedDate = activityData.Item1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Planned Date is not in valid date format. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "ACTUALDATE":
                            activityData = CellData("ACTUALDATE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                DateTime tempDate;
                                bool isDate = DateTime.TryParse(activityData.Item1, out tempDate);
                                if (isDate)
                                {
                                    activity.ActualDate = activityData.Item1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Actual Date is not in valid date format. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "PLANNEDLOADDATE":
                            activityData = CellData("PLANNEDLOADDATE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                DateTime tempDate;
                                bool isDate = DateTime.TryParse(activityData.Item1, out tempDate);
                                if (isDate)
                                {
                                    activity.PlannedLoadDate = activityData.Item1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Planned Load Date is not in valid date format. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "ACTUALLOADDATE":
                            activityData = CellData("ACTUALLOADDATE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                DateTime tempDate;
                                bool isDate = DateTime.TryParse(activityData.Item1, out tempDate);
                                if (isDate)
                                {
                                    activity.ActualLoadDate = activityData.Item1;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Actual Load Date is not in valid date format. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST1":
                            activityData = CellData("SCAFFOLDACTIVITYLIST1", ws, row, importDict);
                            if (activityData != null)
                            {
                                if (activityData.Item1 == "")
                                {
                                    activity.ActivityJobList8ID = Guid.Empty;
                                }
                                else
                                {
                                    ScaffoldTagJobListItem listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID,
                                        activityData.Item1, ListNumberTypes.ActivityList1);
                                    if (listItem.Name.ToUpper() == activityData.Item1.ToUpper())
                                    {
                                        activity.ActivityJobList8ID = listItem.ScaffoldTagJobListItemID;
                                    }
                                    else
                                    {
                                        WriteCell(row, resultCol, ws, "Scaffold Activity List1 value is not valid. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                    }
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST2":
                            activityData = CellData("SCAFFOLDACTIVITYLIST2", ws, row, importDict);
                            if (activityData != null)
                            {
                                if (activityData.Item1 == "")
                                {
                                    activity.ActivityJobList9ID = Guid.Empty;
                                }
                                else
                                {
                                    ScaffoldTagJobListItem listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID,
                                        activityData.Item1, ListNumberTypes.ActivityList2);
                                    if (listItem.Name.ToUpper() == activityData.Item1.ToUpper())
                                    {
                                        activity.ActivityJobList9ID = listItem.ScaffoldTagJobListItemID;
                                    }
                                    else
                                    {
                                        WriteCell(row, resultCol, ws, "Scaffold Activity List2 value is not valid. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                    }
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST3":
                            activityData = CellData("SCAFFOLDACTIVITYLIST3", ws, row, importDict);
                            if (activityData != null)
                            {
                                if (activityData.Item1 == "")
                                {
                                    activity.ActivityJobList10ID = Guid.Empty;
                                }
                                else
                                {
                                    ScaffoldTagJobListItem listItem = ScaffoldTagJobListItem.GetScaffoldTagJobListItem(jobsite.StockingLocationID,
                                        activityData.Item1, ListNumberTypes.ActivityList3);
                                    if (listItem.Name.ToUpper() == activityData.Item1.ToUpper())
                                    {
                                        activity.ActivityJobList10ID = listItem.ScaffoldTagJobListItemID;
                                    }
                                    else
                                    {
                                        WriteCell(row, resultCol, ws, "Scaffold Activity List3 value is not valid. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                    }
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYTEXT":
                            activityData = CellData("SCAFFOLDACTIVITYTEXT", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                activity.ActivityJobText1 = activityData.Item1;
                            }
                            break;
                        case "PLANNEDHOURS":
                            activityData = CellData("PLANNEDHOURS", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDouble = double.TryParse(activityData.Item1, out double tempDouble);
                                if (isDouble)
                                {
                                    activity.PlannedHours = tempDouble;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Planned Hours is not a valid number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "PLANNEDCOST":
                            activityData = CellData("PLANNEDCOST", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDecimal = decimal.TryParse(activityData.Item1, out decimal tempDecimal);
                                if (isDecimal)
                                {
                                    activity.PlannedCost = tempDecimal;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Planned Cost is not a valid number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "ACTUALHOURS":
                            activityData = CellData("ACTUALHOURS", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDouble = double.TryParse(activityData.Item1, out double tempDouble);
                                if (isDouble)
                                {
                                    activity.ActualHours = tempDouble;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Actual Hours is not a valid number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "ACTUALCOST":
                            activityData = CellData("ACTUALCOST", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDecimal = decimal.TryParse(activityData.Item1, out decimal tempDecimal);
                                if (isDecimal)
                                {
                                    activity.ActualCost = tempDecimal;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Actual Cost is not a valid number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;

                    }
                }
            }

            if (importFailure)
            {
                ((Excel.Range)ws.Cells[row, summaryColumn]).Interior.Color = warningColor;
            }
            else
            {
                Tuple<object, object, object, int, int, object>
                    importObject = new Tuple<object, object, object, int, int, object>(activity, null, null, row, resultCol, null);
                importedObjects.Add(importObject);
                WriteCell(row, resultCol, ws, Validated);
            }
            return importFailure;
        }

        private bool ImportUnitOfMeasure(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            HashSet<string> foundNames)
        {
            bool importFailure = false;
            Uom unitOfMeasure = null;
            string description = "";

            Tuple<string, ImportConfig.ImportColunnType, int> activityData;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "NAME":
                            activityData = CellData("NAME", ws, row, importDict);
                            if (activityData != null)
                            {
                                string name = activityData.Item1.ToUpper();

                                if (name == "")
                                {
                                    WriteCell(row, resultCol, ws, "Name can not be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else if (foundNames.Contains(name))
                                {
                                    WriteCell(row, resultCol, ws, "Name can not be duplicated on this sheet. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    unitOfMeasure = Uom.GetUom(name);
                                    if (unitOfMeasure.Name.ToUpper() == name)
                                    {
                                        WriteCell(row, resultCol, ws, "Unit Of Measure Name exists in the database. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        unitOfMeasure = Uom.NewUom();
                                        unitOfMeasure.Name = name;
                                        foundNames.Add(name);
                                    }
                                }
                            }
                            break;
                        case "DESCRIPTION":
                            activityData = CellData("DESCRIPTION", ws, row, importDict);
                            if (activityData != null)
                            {
                                description = activityData.Item1;
                                if (description != "")
                                {
                                    unitOfMeasure.Description = description;
                                }
                                else
                                {
                                    WriteCell(row, resultCol, ws, "Description can't be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                }
                            }
                            break;
                        case "ISACTIVE":
                            activityData = CellData("ISACTIVE", ws, row, importDict);

                            BoolResult result = CellBool(activityData, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "IsActive is Required; must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                            }
                            else
                            {
                                unitOfMeasure.IsActive = result == BoolResult.True ? true : false;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                Tuple<object, object, object, int, int, object>
                    tuple = new Tuple<object, object, object, int, int, object>(unitOfMeasure, null, null, row, resultCol, null);
                importedObjects.Add(tuple);
                WriteCell(row, resultCol, ws, Validated);
            }
            return importFailure;
        }

        private bool ImportUnitHourRateProfile(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects,
            Dictionary<Guid, UnitHourRateProfile> profileDict)
        {
            bool newProfile = false;
            bool importFailure = false;
            StockingLocation jobsite = null;
            UnitHourRate unitHourRate = null;
            UnitHourRateProfile profile = null;

            Tuple<string, ImportConfig.ImportColunnType, int> activityData;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            activityData = CellData("JOBSITENAME", ws, row, importDict);
                            if (activityData != null)
                            {
                                string name = activityData.Item1.ToUpper();
                                if (name == "")
                                {
                                    WriteCell(row, resultCol, ws, "Job Site Name can not be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite = StockingLocation.GetStockingLocation(name, false);
                                    if (jobsite.Name.ToUpper() != name)
                                    {
                                        WriteCell(row, resultCol, ws, "Job Site not found in the database. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    else
                                    {
                                        if (profileDict.ContainsKey(jobsite.StockingLocationID))
                                        {
                                            profile = profileDict[jobsite.StockingLocationID];
                                        }
                                        else
                                        {
                                            profile = UnitHourRateProfile.GetUnitHourRateProfileByJobSite(jobsite.StockingLocationID, false);
                                            if (profile.JobSiteID != jobsite.StockingLocationID)
                                            {
                                                profile = UnitHourRateProfile.NewUnitHourRateProfile(jobsite.StockingLocationID);
                                                newProfile = true;

                                            }
                                            profileDict.Add(jobsite.StockingLocationID, profile);
                                        }
                                    }
                                }
                            }
                            break;
                        case "UNITHOURRATEPROFILE":
                            activityData = CellData("UNITHOURRATEPROFILE", ws, row, importDict);
                            if (activityData != null)
                            {
                                string rateName = activityData.Item1.ToUpper();
                                if (rateName == "")
                                {
                                    WriteCell(row, resultCol, ws, "Unit Hour Rate Profile can not be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }

                                else
                                {
                                    UnitHourRateProfile unitHourRateProfile = UnitHourRateProfile.GetUnitHourRateProfileByJobSite(jobsite.StockingLocationID, false);
                                    if (unitHourRateProfile.UnitHourRates != null)
                                    {
                                        bool nameExists = unitHourRateProfile.UnitHourRates.Where(x => x.Name.ToUpper() == rateName).Any();
                                        if (nameExists != false)
                                        {
                                            WriteCell(row, resultCol, ws, "Unit Hour Rate Profile already exists. ", useColor: true, append: true, color: warningColor);
                                            importFailure = true;
                                            breakLoop = true;
                                        }
                                        else
                                        {
                                            unitHourRate = UnitHourRate.NewUnitHourRate(profile.UnitHourRateProfileID);
                                            unitHourRate.Name = activityData.Item1;
                                        }
                                    }
                                    else
                                    {
                                        unitHourRate = UnitHourRate.NewUnitHourRate(profile.UnitHourRateProfileID);
                                        unitHourRate.Name = activityData.Item1;
                                    }
                                }
                            }
                            break;
                        case "ISACTIVE":
                            activityData = CellData("ISACTIVE", ws, row, importDict);

                            BoolResult result = CellBool(activityData, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "IsActive is Required; must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;

                            }
                            else
                            {
                                unitHourRate.IsActive = result == BoolResult.True ? true : false;
                            }
                            break;
                        case "UNITOFMEASURE":
                            activityData = CellData("UNITOFMEASURE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                string uomUpper = activityData.Item1.ToUpper();

                                Uom uom = Uom.GetUom(uomUpper);
                                if (uom.Name.ToUpper() != uomUpper)
                                {
                                    WriteCell(row, resultCol, ws, "The Unit Of Measure is not found in the database. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    unitHourRate.UOMID = uom.UomID;

                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "The Unit Of Measure is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }
                            break;

                        case "UNITFACTOR":
                            activityData = CellData("UNITFACTOR", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDouble = double.TryParse(activityData.Item1, out double tempDouble);
                                if (!isDouble)
                                {
                                    WriteCell(row, resultCol, ws, "Unit Factor is not a number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    unitHourRate.Factor = tempDouble;
                                }
                            }
                            break;
                        case "SPLIT":
                            activityData = CellData("SPLIT", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDouble = double.TryParse(activityData.Item1, out double tempDouble);
                                if (!isDouble)
                                {
                                    WriteCell(row, resultCol, ws, "Split is not a number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    unitHourRate.Split = tempDouble;
                                }
                            }
                            break;
                        case "PRICE":
                            activityData = CellData("PRICE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDecimal = decimal.TryParse(activityData.Item1, out decimal tempDecimal);
                                if (!isDecimal)
                                {
                                    WriteCell(row, resultCol, ws, "Price is not a number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    unitHourRate.Price = tempDecimal;
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                Tuple<object, object, object, int, int, object>
                    tuple = new Tuple<object, object, object, int, int, object>(profile, unitHourRate, newProfile, row, resultCol, null);
                importedObjects.Add(tuple);
                WriteCell(row, resultCol, ws, Validated);
            }
            return importFailure;
        }

        private bool ImportMultiplier(Dictionary<string, ImportConfig.ImportColumn> importDict, Excel.Worksheet ws,
            ImportConfig config, int row, int resultCol, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool importFailure = false;
            StockingLocation jobsite = null;

            UnitHourRateProfile unitHourRateProfile = null;
            Multiplier multiplier = null;

            Tuple<string, ImportConfig.ImportColunnType, int> activityData;

            bool breakLoop = false;
            foreach (var configCol in config.ImportColumns)
            {
                if (!breakLoop)
                {
                    switch (NormStr(configCol.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            activityData = CellData("JOBSITENAME", ws, row, importDict);
                            if (activityData != null)
                            {
                                string name = activityData.Item1.ToUpper();
                                if (name == "")
                                {
                                    WriteCell(row, resultCol, ws, "Job Site Name can not be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    jobsite = StockingLocation.GetStockingLocation(name, false);
                                    if (jobsite.Name.ToUpper() != name)
                                    {
                                        WriteCell(row, resultCol, ws, "Job Site not found in the database. ", useColor: true, append: true, color: warningColor);
                                        importFailure = true;
                                        breakLoop = true;
                                    }
                                    unitHourRateProfile = UnitHourRateProfile.GetUnitHourRateProfileByJobSite(jobsite.StockingLocationID, false);
                                }
                            }
                            break;

                        case "NAME":
                            activityData = CellData("NAME", ws, row, importDict);
                            if (activityData != null)
                            {
                                string name = activityData.Item1;
                                if (name == "")
                                {
                                    WriteCell(row, resultCol, ws, "Multiplier Name can not be blank. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else if (unitHourRateProfile.Multipliers != null &&
                                    unitHourRateProfile.Multipliers.Where(x => x.Name.ToUpper() == name.ToUpper()).Any())
                                {
                                    WriteCell(row, resultCol, ws, "Multiplier already exists in the database. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    multiplier = Multiplier.NewMultiplier(unitHourRateProfile.UnitHourRateProfileID);
                                    multiplier.Name = name;
                                }
                            }
                            break;

                        case "VALUE":
                            activityData = CellData("VALUE", ws, row, importDict);
                            if (activityData != null && activityData.Item1 != "")
                            {
                                bool isDouble = double.TryParse(activityData.Item1, out double tempDouble);
                                if (!isDouble)
                                {
                                    WriteCell(row, resultCol, ws, "Value is not a number. ", useColor: true, append: true, color: warningColor);
                                    importFailure = true;
                                    breakLoop = true;
                                }
                                else
                                {
                                    multiplier.Value = tempDouble;
                                }
                            }
                            else
                            {
                                WriteCell(row, resultCol, ws, "Value is required. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;
                                breakLoop = true;
                            }

                            break;
                        case "ISACTIVE":
                            activityData = CellData("ISACTIVE", ws, row, importDict);

                            BoolResult result = CellBool(activityData, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                WriteCell(row, resultCol, ws, "IsActive is Required; must be TRUE, FALSE, YES, or NO .  The value is invalid. ", useColor: true, append: true, color: warningColor);
                                importFailure = true;

                            }
                            else
                            {
                                multiplier.IsActive = result == BoolResult.True ? true : false;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }

            if (!importFailure)
            {
                Tuple<object, object, object, int, int, object>
                    tuple = new Tuple<object, object, object, int, int, object>(multiplier, unitHourRateProfile, null, row, resultCol, null);
                importedObjects.Add(tuple);
                WriteCell(row, resultCol, ws, Validated);
            }
            return importFailure;
        }


        /// <summary>
        /// The various imports save data in a Tuple of objects until the import worksheet is validated. 
        /// SaveImportedData saves these objects.  This method does not apply to Balance imiports. 
        /// </summary>
        /// <param name="importIdentifier"></param>
        /// <param name="importedObjects"></param>
        /// <returns></returns>
        private bool SaveImportedData(string importIdentifier, Excel.Worksheet ws, List<Tuple<object, object, object, int, int, object>> importedObjects)
        {
            bool success = true;
            int importIndex = 0;
            switch (NormStr(importIdentifier))
            {
                case "IMPORT:GLOBALLISTS":
                    foreach (var item in importedObjects)
                    {
                        string itemName = "";
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        try
                        {
                            switch (item.Item2)
                            {
                                case GlobalKeyType.Project:
                                    itemName = ((Project)item.Item1).Name;
                                    ((Project)item.Item1).Save();
                                    break;

                                case GlobalKeyType.Priority:
                                    itemName = ((Priority)item.Item1).Name;
                                    ((Priority)item.Item1).Save();
                                    break;

                                case GlobalKeyType.Step:
                                    itemName = ((ScaffoldTagStep)item.Item1).Name;
                                    ((ScaffoldTagStep)item.Item1).Save();
                                    break;

                                case GlobalKeyType.Driver:
                                    itemName = ((Driver)item.Item1).Name;
                                    ((Driver)item.Item1).Save();
                                    break;

                                case GlobalKeyType.Shift:
                                    itemName = ((Shift)item.Item1).Name;
                                    ((Shift)item.Item1).Save();
                                    break;

                                case GlobalKeyType.State:
                                    itemName = ((State)item.Item1).Name;
                                    ((State)item.Item1).Save();
                                    break;

                                case GlobalKeyType.CountyParish:
                                    itemName = ((County)item.Item1).Name;
                                    ((County)item.Item1).Save();
                                    break;
                                case GlobalKeyType.Manufacturer:
                                    itemName = ((Manufacturer)item.Item1).Name;
                                    ((Manufacturer)item.Item1).Save();
                                    break;
                                case GlobalKeyType.RequestList1:
                                case GlobalKeyType.RequestList2:
                                case GlobalKeyType.RequestList3:
                                    itemName = ((GlobalRequestListItem)item.Item1).Name;
                                    ((GlobalRequestListItem)item.Item1).Save();
                                    break;

                                case GlobalKeyType.JobList1:
                                    itemName = ((JobList1)item.Item1).Name;
                                    ((JobList1)item.Item1).Save();
                                    break;

                                case GlobalKeyType.JobList2:
                                    itemName = ((JobList2)item.Item1).Name;
                                    ((JobList2)item.Item1).Save();
                                    break;

                                case GlobalKeyType.ShipmentList1:
                                    itemName = ((GlobalList1)item.Item1).Name;
                                    ((GlobalList1)item.Item1).Save();
                                    break;

                                case GlobalKeyType.ShipmentList2:
                                    itemName = ((GlobalList2)item.Item1).Name;
                                    ((GlobalList2)item.Item1).Save();
                                    break;
                            }
                            WriteCell(item.Item4, item.Item5, ws, Imported);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("duplicate key"))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Unable to save. This entry exists in the database.  "),
                                   useColor: true, append: false, color: warningColor);
                            }
                            else
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Unable to save {0}: value = {1}. Error = {2}. ",
                                    item.Item2.ToString(), itemName, ex.Message),
                                    useColor: true, append: false, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:TAXRATE":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        TaxRate taxRate = (TaxRate)item.Item1;
                        if (taxRate.IsSavable)
                        {
                            try
                            {
                                taxRate.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_TaxRate_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Tax Rate name: Tax Rate Name must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in taxRate.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:PRODUCTCATEGORY":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        ProductCategory category = (ProductCategory)item.Item1;
                        if (category.IsSavable)
                        {
                            try
                            {
                                category.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_ProductCategory_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Product Category name: Product Category Name must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in category.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:PRODUCTCATALOG":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        Product product = (Product)item.Item1;
                        if (product.IsSavable)
                        {
                            try
                            {
                                product.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_ProductCategory_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Part Number : Product Part Number must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in product.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:CONSUMABLESCATEGORY":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        ProductCategory category = (ProductCategory)item.Item1;
                        if (category.IsSavable)
                        {
                            try
                            {
                                category.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_ProductCategory_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Category name: Consumable Category Name must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in category.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:CONSUMABLESCATALOG":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        Product product = (Product)item.Item1;
                        if (product.IsSavable)
                        {
                            try
                            {
                                product.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_ProductCategory_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Part Number : Product Part Number must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in product.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:RENTADJUSTMENT":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        RentalPeriodAdjustment adjust = (RentalPeriodAdjustment)item.Item1;
                        if (adjust.IsSavable)
                        {
                            try
                            {
                                adjust.Save();
                                //WriteCell(item.Item4, item.Item5, ws, Imported);
                                var goodRules = (Dictionary<string, List<int>>)item.Item6;
                                var rowList = goodRules[adjust.Name.ToUpper()];
                                foreach (var rowNumber in rowList)
                                {
                                    WriteCell(rowNumber, item.Item5, ws, Imported);
                                }
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_RentalPeriodAdjustment_Name"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate Rent Period Adjustment Name : Rent Period Adjustment Name must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in adjust.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:RATEPROFILE":
                    Dictionary<string, bool> profilesSaved = new Dictionary<string, bool>();

                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        RateProfile profile = (RateProfile)item.Item1;
                        if (profile.IsSavable)
                        {
                            if (!profilesSaved.ContainsKey(profile.Name))
                            {
                                profile = (RateProfile)importedObjects.Where(x => ((RateProfile)x.Item1).Name == profile.Name).Last().Item1;

                                try
                                {
                                    profile.Save();
                                    profilesSaved.Add(profile.Name, true);
                                    WriteCell(item.Item4, item.Item5, ws, Imported);
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message.Contains("UQ_RateProfile_Name"))
                                    {
                                        WriteCell(item.Item4, item.Item5, ws, "Dupliate Rate Profile Name : Rate Profile Name must be unique.", useColor: true, color: warningColor);
                                    }
                                    else
                                    {
                                        WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                    }
                                }
                            }
                            else
                            {
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in profile.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:CUSTOMER":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        BusinessPartner customer = (BusinessPartner)item.Item1;
                        if (customer.IsSavable)
                        {
                            try
                            {
                                // If the Customer Number is not provided in the data
                                if (!(bool)item.Item2)
                                {
                                    IDNumber number = IDNumber.GetIncrementedIdNumber(IDNumberTypes.Customer);
                                    customer.PartnerNumber = number.Number;
                                }
                                customer.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("Cannot insert duplicate key row in object 'dbo.BusinessPartner' with unique index 'UQ_PartnerNumber'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "The Customer's Number must be unique.  This one is a duplicate. Please correct and retry", append: false, useColor: true, color: warningColor);
                                }
                                else if (ex.Message.Contains("Cannot insert duplicate key row in object"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "This customer is a duplicate. Please correct and retry", append: false, useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in customer.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:CUSTOMERCONTACT":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        AvontusUser user = (AvontusUser)item.Item1;
                        if (user.IsSavable)
                        {
                            try
                            {
                                user.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_AvontusUser_Username_Unique") &&
                                    ex.Message.Contains("Cannot insert duplicate key in object 'dbo.AvontusUser'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws,
                                        string.Format("Contacts can only be used for a single customer / vendor.Username {0} has already been used.", item.Item5),
                                        useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in user.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:ORDER":
                    foreach (var item in importedObjects)
                    {
                        bool savesDone = true;
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        if (item.Item1 != null)
                        {
                            try
                            {
                                Order order = (Order)item.Item1;
                                order.Save();
                            }
                            catch (Exception ex)
                            {
                                savesDone = false;
                                WriteCell(item.Item4, item.Item5, ws, "Save Order Failed: " + ex.Message, useColor: true, color: warningColor);
                            }
                        }
                        if (item.Item2 != null && savesDone)
                        {
                            try
                            {
                                Order WorkOrder = (Order)item.Item2;
                                WorkOrder.Save();
                            }
                            catch (Exception ex)
                            {
                                savesDone = false;
                                WriteCell(item.Item4, item.Item5, ws, "Save WorkOrder Failed: " + ex.Message, useColor: true, color: warningColor);
                            }
                        }

                        if (savesDone)
                        {
                            WriteCell(item.Item4, item.Item5, ws, Imported);
                        }
                    }
                    break;
                case "IMPORT:VENDOR":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        BusinessPartner vendor = (BusinessPartner)item.Item1;
                        if (vendor.IsSavable)
                        {
                            try
                            {
                                // If the Vendor Number is not provided in the data
                                if (!(bool)item.Item2)
                                {
                                    IDNumber number = IDNumber.GetIncrementedIdNumber(IDNumberTypes.Vendor);
                                    vendor.PartnerNumber = number.Number;
                                }
                                vendor.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("Cannot insert duplicate key row in object 'dbo.BusinessPartner' with unique index 'UQ_PartnerNumber'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "The Number's Number must be unique.  This one is a duplicate. Please correct and retry", append: false, useColor: true, color: warningColor);
                                }
                                else if (ex.Message.Contains("Cannot insert duplicate key row in object"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "This vendor is a duplicate. Please correct and retry", append: false, useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in vendor.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:VENDORCONTACT":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        AvontusUser user = (AvontusUser)item.Item1;
                        if (user.IsSavable)
                        {
                            try
                            {
                                user.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_AvontusUser_Username_Unique") &&
                                    ex.Message.Contains("Cannot insert duplicate key in object 'dbo.AvontusUser'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws,
                                        string.Format("Contacts can only be used for a single customer/vendor. Username {0} has already been used.", item.Item5),
                                        useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in user.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:JOBSITE":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        StockingLocation jobsite = (StockingLocation)item.Item1;

                        if ((string)item.Item2 != "")
                        {
                            TradingPartner parent = TradingPartner.GetTradingPartner((string)item.Item2);
                            if (parent.Name.ToUpper() == ((string)item.Item2).ToUpper())
                            {
                                jobsite.ParentTradingPartnerID = parent.TradingPartnerID;
                            }
                            else
                            {
                                WriteCell(item.Item4, item.Item5, ws, "Parent Jobsite is not valid.", useColor: true, color: warningColor);
                            }
                        }

                        if (jobsite.IsSavable)
                        {
                            try
                            {
                                if (jobsite.TrackRequests)
                                {
                                    if (jobsite.RequestStatuses.Count == 0)
                                    {
                                        jobsite.RequestStatuses.CreateDefaultStatuses(jobsite.StockingLocationID);
                                    }
                                }
                                jobsite.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("Cannot insert duplicate key row in object 'dbo.TradingPartner' with unique index 'UQ_TradingPartner_Name'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "This jobsite exists in the database.  Unable to impoort a duplciate job site.",
                                        useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message,
                                        useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in jobsite.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true,
                                    useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:REQUEST":
                    foreach (var item in importedObjects)
                    {
                        bool partialSaveFailed = false;

                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);

                        if (!partialSaveFailed)
                        {
                            Request request = (Request)item.Item1;
                            try
                            {
                                request.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("unique index 'UQ_RequestNumber'"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Request Failed: Duplicate Request Number", useColor: true, color: warningColor);
                                    partialSaveFailed = true;
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Request Failed: " + ex.Message, useColor: true, color: warningColor);
                                    partialSaveFailed = true;
                                }
                            }
                        }
                    }

                    break;
                case "IMPORT:SCAFFOLDJOBLIST":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        ScaffoldTagActivityType type = (ScaffoldTagActivityType)item.Item1;
                        if (type != null && type.IsSavable)
                        {
                            try
                            {
                                type.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("Cannot insert duplicate key row"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Scaffold Job List values must be unique.  This list value is a duplicate.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else if (type != null)
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in type.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }

                        ScaffoldTagJobListItem listItem = (ScaffoldTagJobListItem)item.Item2;
                        if (listItem != null && listItem.IsSavable)
                        {
                            try
                            {
                                listItem.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("Cannot insert duplicate key row"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Scaffold Job List values must be unique.  This list value is a duplicate.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else if (listItem != null)
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in listItem.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {

                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;

                case "IMPORT:SCAFFOLDTAG":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);

                        ScaffoldTag tag = (ScaffoldTag)item.Item1;
                        if (tag.IsSavable)
                        {
                            try
                            {
                                if (item.Item2 != null)
                                    ((Order)item.Item2).Save();
                                if (item.Item3 != null)
                                    ((Order)item.Item3).Save();

                                tag.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Save tag failed. Error = {0}. ", ex.Message), useColor: true, append: true, color: warningColor);
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in tag.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:SCAFFOLDTAGACTIVITY":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        ScaffoldTagActivity activity = (ScaffoldTagActivity)item.Item1;
                        if (activity.IsSavable)
                        {
                            try
                            {
                                activity.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("UQ_ScaffoldTagActivity_Summary"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Dupliate activity key: (Jobsite, Tag, Summary) must be unique.", useColor: true, color: warningColor);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Failed: " + ex.Message, useColor: true, color: warningColor);
                                }
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);

                            foreach (var brokenRule in activity.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                case "IMPORT:UNITOFMEASURE":
                    foreach (var item in importedObjects)
                    {
                        ShowProgress(ws, ProgressDisplayType.Row, ++importIndex, importedObjects.Count, versionStr, item.Item5);
                        System.Diagnostics.Debug.WriteLine("Save row {0}", item.Item4);
                        Uom unitOfMeasure = (Uom)item.Item1;
                        if (unitOfMeasure.IsSavable)
                        {
                            try
                            {
                                unitOfMeasure.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                            catch (Exception ex)
                            {
                                string msg = ex.Message;
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in unitOfMeasure.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }

                    break;

                case "IMPORT:UNITHOURRATEPROFILE":
                    Guid saveJob = Guid.Empty;
                    foreach (var item in importedObjects)
                    {
                        UnitHourRateProfile profile = (UnitHourRateProfile)item.Item1;
                        UnitHourRate rate = (UnitHourRate)item.Item2;

                        bool newProfile = (bool)item.Item3;

                        if (rate.IsSavable == false)
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in profile.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                        else
                        {
                            if (newProfile)
                            {
                                if (profile.IsSavable)
                                {
                                    profile.Save();
                                    profile = UnitHourRateProfile.GetUnitHourRateProfile(profile.UnitHourRateProfileID);
                                }
                            }
                            else
                            {
                                profile = UnitHourRateProfile.GetUnitHourRateProfile(profile.UnitHourRateProfileID);
                            }

                            rate.Save();
                            profile.UnitHourRates.Add(rate);
                            profile.Save();
                            WriteCell(item.Item4, item.Item5, ws, Imported);
                        }
                    }

                    break;

                case "IMPORT:MULTIPLIER":

                    foreach (var item in importedObjects)
                    {
                        UnitHourRateProfile profile = (UnitHourRateProfile)item.Item2;
                        Multiplier multiplier = (Multiplier)item.Item1;

                        if (multiplier.IsSavable)
                        {
                            bool multiplierOK = true;
                            profile = UnitHourRateProfile.GetUnitHourRateProfile(profile.UnitHourRateProfileID);
                            try
                            {
                                multiplier.Save();
                            }
                            catch (Exception ex)
                            {
                                multiplierOK = false;
                                if (ex.Message.Contains("duplicate key"))
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Multiplier exists in database", append: false);
                                }
                                else
                                {
                                    WriteCell(item.Item4, item.Item5, ws, "Multiplier import failed. ", append: false);
                                }
                            }

                            if (multiplierOK)
                            {
                                profile.Multipliers.Add(multiplier);
                                profile.Save();
                                WriteCell(item.Item4, item.Item5, ws, Imported);
                            }
                        }
                        else
                        {
                            WriteCell(item.Item4, item.Item5, ws, useColor: true);
                            foreach (var brokenRule in profile.BrokenRulesCollection.Take(maxBrokenRulesReported))
                            {
                                WriteCell(item.Item4, item.Item5, ws, string.Format("Broken Rule: {0}. ", brokenRule.Description), append: true, useColor: true, color: warningColor);
                            }
                        }
                    }
                    break;
                default:
                    success = false;
                    break;
            }
            ClearProgress(ws);
            return success;
        }

        /// <summary>
        /// Import balances to Scaffold Tags and Job Sites creating Shipments
        /// </summary>
        private void ImportBalances(Excel.Worksheet ws, int maxRows, int maxColumns, int resultRow,
            bool tagBalanceImport, Tuple<int, int> rentStart, Tuple<int, int> tagLabel,
            Tuple<int, int> jobLabel, Tuple<int, int> partNumber, Tuple<int, int> description,
            Tuple<int, int> activityLabel, Tuple<int, int> summaryLabel)
        {
            Dictionary<string, Guid> validatedParts = new Dictionary<string, Guid>();
            List<ImportBalance> importBalances = new List<ImportBalance>();
            bool validationErrorsDetected = false;
            bool brokenColumnRules = false;

            HashSet<string> partNumbers = new HashSet<string>();

            for (int row = tagBalanceImport ? 14 : 11; row < resultRow; row++)
            {
                string data = CellData(ws, row, 1);
                if (partNumbers.Contains(data.ToUpper()))
                {
                    WriteCell(resultRow, 2, ws, string.Format("Duplicate part number = {0}. ", data), useColor: true, append: true);
                    brokenColumnRules = true;
                }
                else
                {
                    partNumbers.Add(data.ToUpper());
                }
            }

            // Import tag/job balances that have not already been imported

            int newMaxCols = maxColumns;
            for (int col = maxColumns; col >= 3; col--)
            {
                if (!EmptyCol(ws, col))
                {
                    newMaxCols = col;
                    break;
                }
            }

            maxColumns = newMaxCols;


            for (int col = rentStart.Item2 + 1; col <= maxColumns; col++)
            {

                brokenColumnRules = false;
                // bool isBillable = false;

                ShowProgress(ws, ProgressDisplayType.Column, col, maxColumns, versionStr);

                string resultStr = ((Excel.Range)ws.Cells[resultRow, col]).Formula.ToString();
                if (resultStr != Imported)
                {
                    bool foundQuantities = false;
                    for (int row = tagBalanceImport ? 14 : 11; row < resultRow; row++)
                    {
                        if (CellData(ws, row, col) != "")
                        {
                            foundQuantities = true;
                            break;
                        }
                    }
                    if (!foundQuantities)
                    {
                        WriteCell(resultRow, col, ws, "No Quantities found to import . ", useColor: true, append: true);
                        brokenColumnRules = true;
                    }

                    // Verify Column is an import column
                    if (tagBalanceImport)
                    {
                        if ((string)((Excel.Range)ws.Cells[rentStart.Item1, col]).Text == "" &&
                                (string)((Excel.Range)ws.Cells[jobLabel.Item1, col]).Text == "" &&
                                (string)((Excel.Range)ws.Cells[tagLabel.Item1, col]).Text == "")
                        {
                            continue;
                        }
                    }
                    else
                    {
                        if ((string)((Excel.Range)ws.Cells[rentStart.Item1, col]).Text == "" &&
                                (string)((Excel.Range)ws.Cells[jobLabel.Item1, col]).Text == "")
                        {
                            continue;
                        }
                    }
                    string rentStartDate = "";

                    if ((string)((Excel.Range)ws.Cells[rentStart.Item1, col]).Text != "")
                    {
                        rentStartDate = ((Excel.Range)ws.Cells[rentStart.Item1, col]).Value.ToString();
                    }

                    bool missingEntity = false;

                    if (CellData(ws, jobLabel.Item1, col) == "")
                    {
                        WriteCell(resultRow, col, ws, "Missing Job Site Name. ", useColor: true, append: true);
                        missingEntity = true;
                    }
                    if (tagBalanceImport && CellData(ws, tagLabel.Item1, col) == "")
                    {
                        WriteCell(resultRow, col, ws, "Missing Tag Name. ", useColor: true, append: true);
                        missingEntity = true;
                    }

                    ImportBalance balance = null;
                    string entityName = "";
                    StockingLocation location = null;

                    if (!missingEntity)
                    {
                        entityName = tagBalanceImport
                            ? ((Excel.Range)ws.Cells[jobLabel.Item1, col]).Value.ToString() +
                                ":" + ((Excel.Range)ws.Cells[tagLabel.Item1, col]).Value.ToString()
                            : ((Excel.Range)ws.Cells[jobLabel.Item1, col]).Value.ToString();

                        int entityRow = tagBalanceImport ? tagLabel.Item1 : jobLabel.Item1;

                        location = StockingLocation.GetStockingLocation(entityName, false);
                        if (location.TrackScaffoldTags)
                        {
                            WriteCell(resultRow, col, ws, "Unable to import job balances to a scaffold tracking job site. ", useColor: true, append: true, color: warningColor, dataRow: entityRow);
                            missingEntity = true;
                            brokenColumnRules = true;
                        }

                        if (location == null || entityName.ToUpper() != location.Name.ToUpper())
                        {
                            WriteCell(resultRow, col, ws, "Invalid Job or Tag Name. ", useColor: true, append: true, color: warningColor, dataRow: entityRow);
                            missingEntity = true;
                            brokenColumnRules = true;
                        }
                        else
                        {
                            balance = new ImportBalance(location.StockingLocationID,
                                location.ParentBranchOrLaydown.StockingLocationID, location.ParentTradingPartnerID, rentStartDate, col);
                        }
                    }
                    if (rentStartDate != "")
                    {
                        try
                        {
                            SmartDate rentStartSmart = new SmartDate(rentStartDate);
                        }
                        catch
                        {
                            string msg = "Invalid Rent Start Date. ";
                            WriteCell(resultRow, col, ws, msg, useColor: true, append: true);
                            brokenColumnRules = true;
                        }
                    }
                    else
                    {
                        string msg = "Missing Rent Start Date. ";
                        WriteCell(resultRow, col, ws, msg, useColor: true, append: true);
                        brokenColumnRules = true;
                    }

                    string activityTypeName = "";
                    string summary = "";

                    if (activityLabel != notFound && tagBalanceImport && !missingEntity)
                    // !(location == null || entityName != location.Name))
                    {
                        if ((string)((Excel.Range)ws.Cells[activityLabel.Item1, col]).Text != "")
                        {
                            activityTypeName = CellData(ws, activityLabel.Item1, col);
                            summary = CellData(ws, summaryLabel.Item1, col);

                            // if there's an activity, there must be a summary; if not, don't try to find it. 
                            if (activityTypeName != "" && summary == "")
                            {
                                string msg = "If there is an Activity value, there must be a Summary value. ";
                                WriteCell(resultRow, col, ws, msg, useColor: true, append: true);
                                brokenColumnRules = true;
                            }
                            else
                            {
                                summary = CellData(ws, summaryLabel.Item1, col);
                                List<Guid> jobList = new List<Guid>();
                                jobList.Add(location.ParentTradingPartner.StockingLocationID);
                                ScaffoldTagList tagList = ScaffoldTagList.GetScaffoldTagList(jobList, false);
                                string tagPart = entityName.Split(':')[1];
                                ScaffoldTagListItem theTag = tagList.Where(x => x.Tag == tagPart).Single();
                                ScaffoldTagActivityList list = ScaffoldTagActivityList.GetScaffoldTagActivityList(theTag.ScaffoldTagID, false, false);
                                var items = list.Where(x => x.ScaffoldTagActivityTypeName.ToUpper() == activityTypeName.ToUpper() &&
                                                            x.Summary.ToUpper() == summary.ToUpper());
                                ScaffoldTagActivityListItem item = null;

                                if (items.Any())
                                {
                                    item = items.Single();
                                    balance.ScaffoldTagActivityTypeID = item.ScaffoldTagActivityID;
                                }
                                else
                                {
                                    balance.ScaffoldTagActivityTypeID = Guid.Empty;
                                    WriteCell(resultRow, col, ws, "Unable to locate Activity Type and Summary for this tag.  ", useColor: true, append: true);
                                    brokenColumnRules = true;
                                }
                            }
                        }
                    }

                    Guid orderID = Guid.Empty;
                    string data = CellData(ws, tagBalanceImport ? 12 : 9, col).ToUpper();
                    BoolResult result = CellBool(data, false, true);
                    switch (result)
                    {
                        case BoolResult.Invalid:
                            WriteCell(resultRow, col, ws, "Is Billable must be True, False, Yes, or No.  ", useColor: true, append: true);
                            brokenColumnRules = true; break;
                        case BoolResult.Empty:
                            if (!missingEntity)
                            {
                                balance.IsBillable = location.DefaultShipmentToBillable;
                            }
                            break;
                        default:
                            if (!missingEntity)
                            {
                                balance.IsBillable = result == BoolResult.True;
                            }
                            break;
                    }

                    if (!missingEntity && balance.IsBillable)
                    {
                        balance.OrderID = missingEntity ? Guid.Empty : location.ParentStockingLocation.OrderID;
                    }


                    if (!brokenColumnRules)
                    {
                        for (int row = description.Item1 + 1; row <= resultRow - 1; row++)
                        {
                            string partNumberStr = ((Excel.Range)ws.Cells[row, partNumber.Item2]).Formula.ToString();

                            // If the sheet has no quantity for a Job/Tag and a part number, then don't import it. 
                            if (((Excel.Range)ws.Cells[row, col]).Formula.ToString() != "")
                            {
                                if (!double.TryParse(((Excel.Range)ws.Cells[row, col]).Formula.ToString(), out double quantity))

                                {
                                    continue;
                                }
                                else
                                {
                                    Guid productID = Guid.Empty;
                                    if (validatedParts.Keys.Contains(partNumberStr))
                                    {
                                        productID = validatedParts[partNumberStr];
                                    }
                                    else
                                    {
                                        Product product = Product.GetProduct(partNumberStr, ProductType.ProductOrConsumable);
                                        if (product.PartNumber == partNumberStr)
                                        {
                                            productID = product.ProductID;
                                            validatedParts.Add(partNumberStr, productID);
                                        }
                                        else
                                        {
                                            string msg = string.Format("Part Number, {0},  is not valid. ", partNumberStr);
                                            WriteCell(resultRow, col, ws, msg, useColor: true, append: true);
                                            brokenColumnRules = true;
                                        }
                                    }
                                    if (Math.Round(quantity, 4) != 0.0)
                                    {
                                        balance.Add(productID, quantity);
                                    }
                                }
                            }
                        }
                    }
                    if (brokenColumnRules)
                    {
                        validationErrorsDetected = true;
                    }
                    else
                    {
                        importBalances.Add(balance);
                        WriteCell(resultRow, col, ws, Validated, append: false);
                    }
                }
            }
            ClearProgress(ws);


            DialogResult response = MessageBoxHelper.Show(string.Format("Validation completed. {0} Do you want to save the Imported items?",
                validationErrorsDetected ? "Validation errors were found as noted." : "No validation errors noted."),
                "Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (response == DialogResult.Yes)
            {
                int shipmentCount = 0;

                foreach (ImportBalance importBalance in importBalances)
                {
                    ShowProgress(ws, ProgressDisplayType.Shipment, ++shipmentCount, importBalances.Count, versionStr);
                    Shipment ship = Shipment.NewShipment(ShipmentStatusType.NewDirectShip);
                    StockingLocation location = StockingLocation.GetStockingLocation(importBalance.StockingLocationID, false);


                    ship.FromStockingLocationID = importBalance.ParentLocationID;
                    ship.ToStockingLocationID = importBalance.StockingLocationID;
                    ship.ShipmentProducts.AddCatalogParts(importBalance.ParentLocationID, importBalance.StockingLocationID, Guid.Empty);
                    ship.RentStartDate = importBalance.RentStartDate;
                    ship.ActualShipDate = importBalance.RentStartDate;  // fix bug 14829
                    ship.DeliveryScaffoldTagActivityID = importBalance.ScaffoldTagActivityTypeID;
                    ship.OrderID = importBalance.OrderID;
                    ship.IsBillable = importBalance.IsBillable;


                    foreach (var balance in importBalance.Balances)
                    {
                        bool found = false;
                        foreach (ShipmentProduct shipProduct in ship.ShipmentProducts)
                        {
                            if (shipProduct.BaseProductID == balance.Item1)
                            {
                                double qty = qty = shipProduct.SentQuantity ?? 0.0;
                                qty += balance.Item2;
                                shipProduct.SentQuantity = qty;
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            string msg = "Unexpected Product failure. ";
                            WriteCell(resultRow, importBalance.WorksheetColumn, ws, msg, useColor: true, append: false, color: warningColor);
                        }
                    }

                    bool shipSaved = true;
                    if (ship.IsSavable)
                    {
                        try
                        {
                            ship.Save();
                        }
                        catch (DataPortalException e)
                        {
                            string msg = string.Format("Unexpected failure saving shiipment. Msg = {0}.", e.Message);
                            WriteCell(resultRow, importBalance.WorksheetColumn, ws, msg, useColor: true, append: false, color: warningColor);
                            shipSaved = false;
                        }
                    }
                    else
                    {
                        string msg;
                        if (ship.BrokenRulesCollection[0].Description == "Activity is required")
                        {
                            msg = "Activity is required; shipment is not savable. ";
                        }
                        else
                        {
                            msg = string.Format("Unable to save shipment; error = {0}. ", ship.BrokenRulesCollection[0].Description);
                        }
                        WriteCell(resultRow, importBalance.WorksheetColumn, ws, msg, useColor: true, append: false, color: warningColor);
                        shipSaved = false;
                    }

                    if (shipSaved)
                    {
                        WriteCell(resultRow, importBalance.WorksheetColumn, ws, Imported, append: false);
                    }
                }
                ClearProgress(ws);

                MessageBoxHelper.Show("Import complete.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Import balances to Branches and Laydown Yards creating Adjustments
        /// </summary>
        private void ImportBranchBalances(Excel.Worksheet ws, int maxRows, int maxColumns, int resultRow)
        {
            Dictionary<string, Guid> validatedParts = new Dictionary<string, Guid>();
            List<BranchImportBalance> importBalances = new List<BranchImportBalance>();
            bool validationErrorsDetected = false;
            bool brokenColumnRules = false;

            int columnsValidated = 0;
            HashSet<string> partNumbers = new HashSet<string>();

            for (int row = 11; row < resultRow; row++)
            {
                string data = CellData(ws, row, 1);
                if (partNumbers.Contains(data.ToUpper()))
                {
                    WriteCell(resultRow, 2, ws, string.Format("Duplicate part number = {0}. ", data), useColor: true, append: true);
                    brokenColumnRules = true;
                }
                else
                {
                    partNumbers.Add(data.ToUpper());
                }
            }

            for (int col = 3; col <= maxColumns; col++)
            {
                ShowProgress(ws, ProgressDisplayType.Column, col, maxColumns - 2, versionStr);

                string resultStr = ((Excel.Range)ws.Cells[resultRow, col]).Formula.ToString();
                if (resultStr != Imported)
                {
                    columnsValidated++;
                    // Verify Column is an import column
                    if (CellData(ws, 7, col) == "" &&
                        CellData(ws, 8, col) == "" &&
                        CellData(ws, 9, col) == "")
                    {
                        continue;
                    }

                    bool foundQuantities = false;
                    for (int row = 11; row < resultRow; row++)
                    {
                        if (CellData(ws, row, col) != "")
                        {
                            foundQuantities = true;
                            break;
                        }
                    }
                    if (!foundQuantities)
                    {
                        WriteCell(resultRow, col, ws, "No Quantities found to import . ", useColor: true, append: true);
                        brokenColumnRules = true;
                    }
                    string name = CellData(ws, 8, col);
                    string branchOrLaydown = CellData(ws, 9, col);
                    int entityRow = 8;  // the name row
                    PartnerTypes partnerType = PartnerTypes.All;

                    switch (branchOrLaydown.ToUpper())
                    {
                        case "BRANCH":
                            partnerType = PartnerTypes.BranchOffice;
                            break;
                        case "LAYDOWN":
                            partnerType = PartnerTypes.LaydownYard;
                            break;
                    }

                    if (partnerType == PartnerTypes.All)
                    {
                        WriteCell(resultRow, col, ws, "Invalid BranchOrLaydown value. ", useColor: true, append: true, color: warningColor, dataRow: entityRow);
                        brokenColumnRules = true;

                    }

                    BranchImportBalance balance = null;
                    StockingLocationList locations = StockingLocationList.GetBranchOfficesAndLaydownYards(false, Guid.Empty);
                    StockingLocationListItem location = null;
                    try
                    {
                        location = locations.Where(x => x.Name.ToUpper() == name.ToUpper() && x.PartnerType == partnerType).Single();
                    }
                    catch
                    {
                        // Pass
                    }

                    if (location == null || location.Name.ToUpper() != name.ToUpper())
                    {
                        WriteCell(resultRow, col, ws, "Invalid Name. ", useColor: true, append: true, color: warningColor, dataRow: entityRow);
                        brokenColumnRules = true;
                    }
                    else
                    {
                        balance = new BranchImportBalance(location.StockingLocationID, col);
                    }

                    if (!brokenColumnRules)
                    {
                        for (int row = 11; row < resultRow; row++)
                        {
                            if (!EmptyRow(ws, row, maxColumns + 1))
                            {
                                string partNumberStr = ((Excel.Range)ws.Cells[row, 1]).Formula.ToString();

                                // If the sheet has no quantity for a part number, then don't import it. 
                                if (((Excel.Range)ws.Cells[row, col]).Formula.ToString() != "")
                                {
                                    Double quantity = 0;
                                    if (!Double.TryParse(((Excel.Range)ws.Cells[row, col]).Formula.ToString(), out quantity))

                                    {
                                        WriteCell(resultRow, col, ws, string.Format("Invalid Quantity in row {0}. ", row), useColor: true, append: true, color: warningColor, dataRow: entityRow);
                                        brokenColumnRules = true;
                                    }
                                    else
                                    {
                                        Guid productID = Guid.Empty;
                                        if (validatedParts.Keys.Contains(partNumberStr))
                                        {
                                            productID = validatedParts[partNumberStr];
                                        }
                                        else
                                        {
                                            Product product = Product.GetProduct(partNumberStr, ProductType.ProductOrConsumable);
                                            if (product.PartNumber == partNumberStr)
                                            {
                                                if (product.IsSerialized)
                                                {
                                                    string msg = string.Format("Part Number, {0},  is a serialized part; it's balance can't be imported. ", partNumberStr);
                                                    WriteCell(resultRow, col, ws, msg, useColor: true, append: true, color: warningColor, dataRow: entityRow, dataCol: col);
                                                    brokenColumnRules = true;
                                                }
                                                else
                                                {
                                                    productID = product.ProductID;
                                                    validatedParts.Add(partNumberStr, productID);
                                                }
                                            }
                                            else
                                            {
                                                if (partNumberStr == "")
                                                {
                                                    partNumberStr = "\"\"";
                                                }
                                                string msg = string.Format("Part Number, {0},  is not valid. ", partNumberStr);
                                                WriteCell(resultRow, col, ws, msg, useColor: true, append: true, color: warningColor, dataRow: entityRow, dataCol: col);
                                                brokenColumnRules = true;
                                            }
                                        }
                                        if (Math.Round(quantity, 4) != 0.0)
                                        {
                                            balance.Add(productID, quantity);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (brokenColumnRules)
                    {
                        validationErrorsDetected = true;
                    }
                    else
                    {
                        importBalances.Add(balance);
                        WriteCell(resultRow, col, ws, Validated, append: false);
                    }
                }
            }
            ClearProgress(ws);
            DialogResult response;
            if (columnsValidated == 0)
            {
                MessageBoxHelper.Show("No Branch or Laydown Balance columns were validated. There is no data to import. ", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;   // EXIT POINT
            }

            response = MessageBoxHelper.Show(string.Format("Validation completed. {0} Do you want to save the Imported items?",
                validationErrorsDetected ? "Validation errors were found as noted." : "No validation errors noted."),
                "Excel Importer", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (response == DialogResult.Yes)
            {
                int columnCount = 0;

                // Dictionary to lookup part numbers encountered
                Dictionary<Guid, string> seenProducts = new Dictionary<Guid, string>();

                int adjustmentCount = 0;
                foreach (BranchImportBalance importBalance in importBalances)
                {
                    ShowProgress(ws, ProgressDisplayType.Column, ++columnCount, maxColumns, versionStr);
                    try
                    {
                        StockedProductAdjustmentCollection adjusts =
                            StockedProductAdjustmentCollection.GetStockedProductAdjustmentCollection(importBalance.StockingLocationID, true, false, ProductType.Product);

                        ShowProgress(ws, ProgressDisplayType.Adjustment, ++adjustmentCount, importBalances.Count, versionStr);

                        // Purge existing balances of imported quantities
                        foreach (Tuple<Guid, double> balance in importBalance.Balances)
                        {
                            Product product;
                            string partNumber = "";
                            if (seenProducts.ContainsKey(balance.Item1))
                            {
                                partNumber = seenProducts[balance.Item1];
                            }
                            else
                            {
                                product = Product.GetProduct(balance.Item1);
                                partNumber = product.PartNumber;
                                seenProducts.Add(balance.Item1, partNumber);
                            }
                            if (adjusts.Contains(partNumber))
                            {
                                if ((adjusts[balance.Item1].QuantityForRent ?? 0.0) != 0.0)
                                {
                                    adjusts[balance.Item1].QuantityForRent = null;
                                }
                            }
                        }

                        StockedProductAdjust.AdjustStockedProducts(adjusts, importBalance.StockingLocationID, "1. Purge initial stock"); // "2. Import balances");

                        adjusts = StockedProductAdjustmentCollection.GetStockedProductAdjustmentCollection(importBalance.StockingLocationID, true, false, ProductType.Product);

                        System.Threading.Thread.Sleep(1200); // delay to force diff3erent time for 2nd adjustment

                        // Import the new balances
                        foreach (Tuple<Guid, double> balance in importBalance.Balances)
                        {
                            Product product;
                            string partNumber = "";
                            if (seenProducts.ContainsKey(balance.Item1))
                            {
                                partNumber = seenProducts[balance.Item1];
                            }
                            else
                            {
                                product = Product.GetProduct(balance.Item1);
                                partNumber = product.PartNumber;
                                seenProducts.Add(balance.Item1, partNumber);
                            }
                            if (adjusts.Contains(partNumber))
                            {

                                adjusts[balance.Item1].QuantityForRent = balance.Item2;
                            }
                        }
                        StockedProductAdjust.AdjustStockedProducts(adjusts, importBalance.StockingLocationID, "2. Import balances");

                        WriteCell(resultRow, importBalance.WorksheetColumn, ws, Imported);
                    }
                    catch (Exception ex)
                    {
                        string msg = string.Format("Unexpected failure saving shipment. Msg = {0}.", ex.Message);
                        WriteCell(resultRow, importBalance.WorksheetColumn, ws, msg, useColor: true, append: true, color: warningColor);
                    }
                }

                ClearProgress(ws);

                MessageBoxHelper.Show("Import complete.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Non Standard Imports
        /// <summary>
        /// This method written to import balances for a specific customer (Downer) to meet load data from 
        /// a non-standard worksheet. 
        /// 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="maxRows"></param>
        /// <param name="maxColumns"></param>
        /// <returns></returns>
        private bool NonStandardImport1(Excel.Worksheet ws, int maxRows, int maxColumns)
        {
            const int jobCol = 1;
            const int tagCol = 2;
            const int activityCol = 3;
            const int partCol = 4;
            const int descrCol = 5;
            const int quantCol = 6;
            const int resultCol = 8;
            string saveJob = "";
            string saveTag = "";

            StockingLocation location = null;
            Shipment ship = null;
            bool skipThisTag = false;

            for (int row = 3; row < maxRows; row++)
            {
                double quantity = double.Parse(((Excel.Range)ws.Cells[row, quantCol]).Value.ToString());
                if ((Excel.Range)ws.Cells[row, jobCol] != null && quantity != 0.0)
                {
                    string job = ((Excel.Range)ws.Cells[row, jobCol]).Value.ToString();
                    string tag = ((Excel.Range)ws.Cells[row, tagCol]).Value.ToString();
                    string activityName = ((Excel.Range)ws.Cells[row, activityCol]).Value.ToString();

                    if (job != saveJob || tag != saveTag)
                    {
                        string name = job + ':' + tag;
                        location = StockingLocation.GetStockingLocation(name, false);
                        if (location.Name != name)
                        {
                            WriteCell(row, resultCol, ws, "Invalid Tag. ");
                            continue;
                        }
                        // Save shipment
                        if (ship != null && ship.IsSavable)
                        {
                            try
                            {
                                ship.Save();
                            }
                            catch (Exception ex)
                            {
                                string msg = ex.Message;
                            }
                            Debug.WriteLine(string.Format("Job {0}, Tag {1}, Shipment {2}, Row {3}", job, tag, ship.ShipmentNumber, row));
                        }
                        else if (ship != null)
                        {
                            // TODO:  Ought to have a warning here, I guess?
                        }

                        // Prepare new shipment
                        saveJob = job;
                        saveTag = tag;
                        skipThisTag = false;

                        ship = Shipment.NewShipment(ShipmentStatusType.NewDirectShip);
                        List<Guid> jobList = new List<Guid>();
                        jobList.Add(location.ParentTradingPartner.StockingLocationID);
                        ScaffoldTagList tagList = ScaffoldTagList.GetScaffoldTagList(jobList, false);
                        ScaffoldTagListItem theTag = tagList.Where(x => x.Tag == tag).Single();
                        ScaffoldTagActivityList list = ScaffoldTagActivityList.GetScaffoldTagActivityList(theTag.ScaffoldTagID, false, false);
                        var items = list.Where(x => x.ScaffoldTagActivityTypeName.ToUpper() == activityName.ToUpper());
                        ScaffoldTagActivityListItem item = null;
                        if (items.Any())
                        {
                            item = items.Single();
                            ship.DeliveryScaffoldTagActivityID = item.ScaffoldTagActivityID;
                        }
                        else
                        {
                            WriteCell(row, resultCol, ws, "Invalid Actity");
                            job = "";
                            tag = "";
                            skipThisTag = true;
                            continue;
                        }

                        ship.FromStockingLocationID = location.ParentBranchOrLaydownOrStagingArea.StockingLocationID;
                        ship.ToStockingLocationID = location.StockingLocationID;
                        ship.ShipmentProducts.AddCatalogParts(ship.FromStockingLocationID, location.StockingLocationID, Guid.Empty);
                        ship.RentStartDate = DateTime.Today.ToShortDateString();
                    }

                    if (skipThisTag)
                    {
                        WriteCell(row, resultCol, ws, "Invalid Actity");
                        continue;
                    }

                    string part = ((Excel.Range)ws.Cells[row, partCol]).Value.ToString();
                    string descr = ((Excel.Range)ws.Cells[row, descrCol]).Value.ToString();

                    if (quantity == 0.0)
                    {
                        WriteCell(row, resultCol, ws, "Zero Quantity");
                    }
                    else
                    {
                        bool found = false;
                        foreach (ShipmentProduct shipProduct in ship.ShipmentProducts)
                        {
                            if (shipProduct.PartNumber == part)
                            {
                                double qty = qty = shipProduct.SentQuantity ?? 0.0;
                                qty += quantity;
                                shipProduct.SentQuantity = qty;
                                WriteCell(row, resultCol, ws, "Done.");
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            WriteCell(row, resultCol, ws, "Product Not Found");
                        }
                    }
                }
            }

            MessageBoxHelper.Show("Import complete.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

        #endregion

        #region Workbook Validation
        internal void WorkbookValidation(string versionStr)
        {

            // If workbook or sheet is null then exit
            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null ||
                Globals.QuantifyImporter.Application.ActiveSheet == null)
            {
                return;
            }

            Excel.Worksheet savedSheet = ((Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet);

            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null)
            {
                MessageBoxHelper.Show("This workbook is currently in protected mode. Please enable editing to continue.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Excel.Workbook wb = Globals.QuantifyImporter.Application.ActiveWorkbook;


            // Create a new Validation Sheet and list of Sheet Names
            Dictionary<string, string> sheetDic = new Dictionary<string, string>();

            List<string> sheetNames = new List<string>();
            foreach (Excel.Worksheet sh in wb.Sheets)
            {
                sheetNames.Add(sh.Name);
            }

            // Create Header for Validation Sheet
            Excel.Worksheet validationSheet = GetValidationSheet(sheetNames, wb);
            WriteCell(1, 1, validationSheet, "Workbook Validation");
            WriteCell(2, 1, validationSheet, "Excel Importer Workbook");
            WriteCell(3, 1, validationSheet, string.Format("Template vrs {0} ", versionStr));
            WriteCell(5, 1, validationSheet, DateTime.Now.ToString(), numberformat: "dd-MMM-yyyy HH:mm");
            WriteCell(7, 1, validationSheet, "Worksheets Found");

            int validationRow = 8;

            var configSeq = ImportConfig.GetAllImportConfigs().OrderBy(x => x.SheetOrder);

            // Print sheets found with import type identified
            foreach (string name in sheetNames)
            {
                WriteCell(validationRow, 2, validationSheet, name);
                if (name.ToUpper() == validationSheet.Name.ToUpper())
                {
                    WriteCell(validationRow, 3, validationSheet, "This Sheet");
                }
                else
                {
                    string sheetIdentifier = GetSheetIdentifier(name, wb, configSeq);
                    if (sheetIdentifier != NonImportSheetID)
                    {
                        if (!sheetDic.ContainsKey(NormStr(sheetIdentifier)))
                        {
                            sheetDic.Add(NormStr(sheetIdentifier), name);
                        }
                        else
                        {
                            sheetIdentifier = string.Format("Duplicate sheet not validated ({0}).", sheetIdentifier);
                        }
                    }
                    WriteCell(validationRow, 3, validationSheet, sheetIdentifier);
                    validationRow += 1;
                }
            }
            validationRow += 3;

            // List import types not represented in the sheets of the workbook
            bool foundMissingSheet = false;
            foreach (ImportConfig config in configSeq)
            {
                if (!sheetDic.ContainsKey(NormStr(config.ImportIdentifier)))
                {
                    if (!foundMissingSheet)
                    {
                        WriteCell(validationRow++, 1, validationSheet, "Missing Import Sheets");
                        foundMissingSheet = true;
                    }
                    WriteCell(validationRow, 2, validationSheet, config.ImportIdentifier);
                    WriteCell(validationRow++, 3, validationSheet, "Import Sheet not found. ");
                }
            }

            if (!sheetDic.ContainsKey("IMPORT:BRANCHORLAYDOWNBALANCES"))
            {
                WriteCell(validationRow, 2, validationSheet, "Import: Branch or Laydown Balances");
                WriteCell(validationRow++, 3, validationSheet, "Import Sheet not found. ");
            }
            if (!sheetDic.ContainsKey("IMPORT:SCAFFOLDTAGBALANCES"))
            {
                WriteCell(validationRow, 2, validationSheet, "Import: Scaffold Tag Balances");
                WriteCell(validationRow++, 3, validationSheet, "Import Sheet not found. ");
            }

            if (!sheetDic.ContainsKey("IMPORT:JOBSITEBALANCES"))
            {
                WriteCell(validationRow, 2, validationSheet, "Import: Job Site Balances");
                WriteCell(validationRow++, 3, validationSheet, "Import Sheet not found. ");
            }

            foreach (ImportConfig config in configSeq)
            {

                bool foundit = sheetDic.TryGetValue(NormStr(config.ImportIdentifier), out string sheetName);
                if (foundit)
                {
                    Excel.Worksheet ws = GetSheet(wb, sheetName);
                    var range = (object[,])ws.UsedRange.Value;
                    if (range == null)
                    {
                        WriteCell(validationRow, 1, validationSheet, sheetName);
                        WriteCell(validationRow++, 2, validationSheet, "Error: nothing to import");
                        return;
                    }

                    int maxRows = range.GetLength(0);
                    int maxColumns = range.GetLength(1);
                    int resultCol = SetResultColumn(ws, maxColumns, HeaderRow, Results);
                    RemoveColumnBackgroundColors(ws, config, HeaderRow, maxRows, resultCol);

                    if (ws != null)
                    {
                        validationRow += 1;
                        WriteCell(validationRow++, 1, validationSheet, sheetName);
                        CheckWorksheetHeaders(ws, maxRows, maxColumns, config.ImportIdentifier, Results, resultCol,
                            out string headerIssues, out ImportConfig importConfig);
                        if (headerIssues != "")
                        {
                            // if the column headers are wrong, then report that and give up on validating this sheet 
                            WriteCell(validationRow++, 2, validationSheet, headerIssues);
                            continue;
                        }
                        else
                        {
                            switch (NormStr(config.ImportIdentifier))
                            {
                                case "IMPORT:GLOBALLISTS":
                                    ValidateGlobalLists(ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig);
                                    break;
                                case "IMPORT:PRODUCTCATEGORY":
                                    ValidateProductCategory(ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig);
                                    break;
                                case "IMPORT:PRODUCTCATALOG":
                                    ValidateProductCatalog(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:CONSUMABLESCATEGORY":
                                    ValidateConsumablesCategory(ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig);
                                    break;
                                case "IMPORT:CONSUMABLESCATALOG":
                                    ValidateConsumablesCatalog(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:RENTADJUSTMENT":
                                    ValidateRentAdjustment(ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig);
                                    break;
                                case "IMPORT:RATEPROFILE":
                                    ValidateRateProfile(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:CUSTOMER":
                                    ValidateCustomer(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:CUSTOMERCONTACT":
                                    ValidateCustomerContact(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:ORDER":
                                    ValidateOrder(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:VENDOR":
                                    ValidateVendor(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:VENDORCONTACT":
                                    ValidateVendorContact(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:TAXRATE":
                                    ValidateTaxRate(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:JOBSITE":
                                    ValidateJobSite(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:REQUEST":
                                    ValidateRequest(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:SCAFFOLDJOBLIST":
                                    ValidateScaffoldJobList(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:SCAFFOLDTAG":
                                    ValidateScaffoldTag(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:SCAFFOLDTAGACTIVITY":
                                    ValidateScaffoldTagActivity(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;

                                case "IMPORT:UNITOFMEASURE":
                                    ValidateUnitOfMeasure(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:UNITHOURRATEPROFILE":
                                    ValidateUnitHourRateProfile(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;
                                case "IMPORT:MULTIPLIER":
                                    ValidateMultiplier(wb, ws, maxRows, resultCol, ref validationRow, validationSheet, importConfig, sheetDic);
                                    break;

                                default:
                                    WriteCell(validationRow++, 2, validationSheet, "ERROR:  Import Type was not recognized!!!", append: true);
                                    break;
                            }
                        }
                    }
                }
            }

            if (sheetDic.ContainsKey("IMPORT:SCAFFOLDTAGBALANCES"))
            {
                ValidateScaffoldTagBalance(wb, ref validationRow, validationSheet, sheetDic);
            }

            if (sheetDic.ContainsKey("IMPORT:JOBSITEBALANCES"))
            {
                ValidateJobSiteBalance(wb, ref validationRow, validationSheet, sheetDic);
            }

            // Balance Import validations
            if (sheetDic.ContainsKey("IMPORT:BRANCHORLAYDOWNBALANCES"))
            {
                ValidateBranchOrLaydownBalance(wb, ref validationRow, validationSheet, sheetDic);
            }

            if (Debugger.IsAttached)
            {
                // Restore sheet where started, for testing ease
                savedSheet.Activate();
            }
        }

        internal void ValidateGlobalLists(Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig)
        {
            bool foundStates = false;
            bool foundCounties = false;
            string msg = "";
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();

            for (int row = HeaderRow + 1; row < maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }

                string name = "";
                foreach (var column in importConfig.ImportColumns)
                {
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "LISTNAME":
                            name = NormStr(CellData(ws, row, column.Column));
                            if (!ValidGlobalListKeys.Contains(name))
                            {
                                msg = "List name invalid. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            switch (name)
                            {
                                case "STATES/PROVINCES":
                                    foundStates = true;
                                    break;
                                case "COUNTIES/PARISHES":
                                    foundCounties = true;
                                    break;
                            }
                            break;
                        case "VALUE/NAME":
                            string key = name + "|" + CellData(ws, row, column.Column).ToUpper();
                            if (foundEntries.ContainsKey(key))
                            {
                                msg = "This entry is a duplciate. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(key, 0);
                            }
                            break;
                        case "ISACTIVE":
                            var data = CellData(ws, row, column.Column);
                            BoolResult result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
            if (!foundStates)
            {
                WriteCell(validationRow++, 2, validationSheet, "Warning: No entries were found for States / Provinces");
            }
            if (!foundCounties)
            {
                WriteCell(validationRow++, 2, validationSheet, "Warning: No entries were found for Counties / Parishes");
            }
        }

        internal void ValidateProductCategory(Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig)
        {
            string msg;
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "PARENTNAME":
                            if (data != "" && !foundEntries.ContainsKey(data))
                            {
                                msg = "Parent Product Category must be imported before it is referenced. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundEntries.ContainsKey(data))
                            {
                                msg = "This is a duplicate entry. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data, 0);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateProductCatalog(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            string msg;
            Excel.Worksheet categorySheet = FindSheet("IMPORT:PRODUCTCATEGORY", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "PARTNUMBER":
                            if (data == "")
                            {
                                msg = "Part Number is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundEntries.ContainsKey(data))
                            {
                                msg = "This is a duplicate entry. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data, 0);
                            }
                            break;
                        case "CATEGORY":
                            if (data != "")
                            {
                                if (categorySheet == null)
                                {
                                    msg = "Can't validate Category; can't find Product Category import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                                else
                                {
                                    if (!ValidateSheetValue(data, categorySheet, 2))
                                    {
                                        msg = "{0}{1}{2}Product Category import sheet ({3}) doesn't contain Category value. ";
                                        WriteCell(row, resultCol, ws, string.Format(msg, "", "", "", categorySheet.Name), append: true);
                                        WriteCell(validationRow++, 2, validationSheet,
                                            string.Format(msg, "Row ", row, ", ", categorySheet.Name), append: true);
                                    }
                                }
                            }
                            break;
                        case "MANUFACTURER":
                            if (data != "")
                            {
                                if (globalListsSheet == null)
                                {
                                    msg = "Can't validate Manufacturer ; can't find Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet,
                                        string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                                else
                                {
                                    if (!ValidateSheetValue(data, globalListsSheet, 2, "MANUFACTURERS", 1))
                                    {
                                        msg = "{0}{1}{2}Global Lists import sheet ({3}) doesn't contain Manufacturer value. ";
                                        WriteCell(row, resultCol, ws,
                                            string.Format(msg, "", "", "", globalListsSheet.Name), append: true);
                                        WriteCell(validationRow++, 2, validationSheet,
                                            string.Format(msg, "Row ", row, ", ", globalListsSheet.Name), append: true);
                                    }
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateConsumablesCategory(Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig)
        {
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            string msg;

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "PARENTNAME":
                            if (data != "" && !foundEntries.ContainsKey(data))
                            {
                                msg = "Parent Consumables Category must be imported before it is referenced. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundEntries.ContainsKey(data))
                            {
                                msg = "This is a duplicate entry. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data, 0);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateConsumablesCatalog(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            string msg;
            Excel.Worksheet categorySheet = FindSheet("IMPORT:CONSUMABLESCATEGORY", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "PARTNUMBER":
                            if (data == "")
                            {
                                msg = "Part Number is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundEntries.ContainsKey(data))
                            {
                                msg = "This is a duplicate entry. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data, 0);
                            }
                            break;
                        case "CATEGORY":
                            if (data != "")
                            {
                                if (categorySheet == null)
                                {
                                    msg = "Can't validate Category; can't find Consumables Category import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                                else
                                {
                                    if (!ValidateSheetValue(data, categorySheet, 2))
                                    {
                                        msg = "{0}{1}{2}Consumables Category import sheet({3}) doesn't contain Category value. ";
                                        WriteCell(row, resultCol, ws, string.Format(msg, "", "", "", categorySheet.Name), append: true);
                                        WriteCell(validationRow++, 2, validationSheet,
                                            string.Format(msg, "Row ", row, ", ", categorySheet.Name), append: true);
                                    }
                                }
                            }
                            break;
                        case "MANUFACTURER":
                            if (data != "")
                            {
                                if (globalListsSheet == null)
                                {
                                    msg = "Can't validate Manufacturer ; can't find Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                                else
                                {
                                    if (!ValidateSheetValue(data, globalListsSheet, 2, "MANUFACTURERS", 1))
                                    {
                                        msg = "{0}{1}{2}Global Lists import sheet ({0}) doesn't contain Manufacturer value. ";
                                        WriteCell(row, resultCol, ws, string.Format(msg, "", "", "", categorySheet.Name), append: true);
                                        WriteCell(validationRow++, 2, validationSheet,
                                            string.Format(msg, "Row ", row, ", ", categorySheet.Name), append: true);
                                    }
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateRentAdjustment(Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig)
        {
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            string msg;
            string name = "";
            string dateOrDays = "";

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                bool flaggedStart = false;
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            name = data;
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "RULENAME":
                            string ruleKey = name + "|" + data;
                            if (data == "")
                            {
                                msg = "Rule Name is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundEntries.ContainsKey(ruleKey))
                            {
                                msg = "This is a duplicate entry for Name and Rule Name. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(ruleKey, 0);
                            }
                            break;
                        case "DATEORDAYS":

                            if (data != "DATE" && data != "DAYS")
                            {
                                msg = "Date or Days must be 'Date' or 'Days'; it is required. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            dateOrDays = data;
                            break;
                        case "START":
                        case "END":
                            switch (dateOrDays)
                            {
                                case "DATE":
                                    DateTime tempDate;
                                    bool dateOK = DateTime.TryParse(data, out tempDate);
                                    if (!dateOK && !flaggedStart)
                                    {
                                        msg = "Start and End are required and must be valid dates. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);

                                        flaggedStart = true;
                                    }
                                    break;
                                case "Days":
                                    int tempInt;
                                    bool intOK = int.TryParse(data, out tempInt);
                                    if (!intOK && !flaggedStart)
                                    {

                                        msg = "Start and End are required and must be integers. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                        flaggedStart = true;
                                    }
                                    break;
                                case "":

                                    msg = "Start and End are required. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                    break;
                            }
                            break;
                        case "DISCOUNTORMARKUP":

                            if (data != "DISCOUNT" && data != "MARKUP")
                            {
                                msg = "Discount or Markup must be 'Discount' or 'Markup'; it is required. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "VALUE":
                            double tempDouble;
                            bool doubleOk = double.TryParse(data, out tempDouble);
                            if (!doubleOk)
                            {
                                msg = "Value is required and must be a number. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateRateProfile(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, bool> isActiveDic = new Dictionary<string, bool>();

            string name = "";

            Excel.Worksheet productCatalogSheet = FindSheet("IMPORT:PRODUCTCATALOG", wb, sheetDic);
            Excel.Worksheet consumableCatalogSheet = FindSheet("IMPORT:CONSUMABLESCATALOG", wb, sheetDic);
            Excel.Worksheet adjustmentSheet = FindSheet("IMPORT:RENTADJUSTMENT", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            name = data;
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (isActiveDic.ContainsKey(name))
                            {
                                if (!isActiveDic[name] == resultTrue)
                                {
                                    msg = "Is Active must be the same for rows with the same Name. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                isActiveDic.Add(name, resultTrue);
                            }

                            break;
                        case "LABELFORHEADER":
                            if (data == "")
                            {
                                msg = "Label For Header is a required value; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "RENTALDAYS":
                            int tempInt;
                            bool isInt = int.TryParse(data, out tempInt);
                            if (!isInt || tempInt < 1)
                            {
                                msg = "Rental Days must be an integer greater than zero. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "PARTNUMBER":      // TODO: Work in progress - adding consumables and recurring charges
                            if (data == "")
                            {
                                msg = "Part Number is required. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                bool found = ValidateSheetValue(data, productCatalogSheet, 1) ||
                                             ValidateSheetValue(data, consumableCatalogSheet, 1);
                                if (!found)
                                {
                                    msg = "Part Number not found in Product Catalog or Consumable import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }

                            string namePart = name + "|" + data;
                            if (foundEntries.ContainsKey(namePart))
                            {
                                msg = "Duplicate [Name, PartNumber]. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(namePart, 0);
                            }

                            break;
                        case "RATE":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Rate must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "RENTQUANTITY":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Rent Quantity must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "JOBCOST":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Job Cost must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "DELIVERYCHARGE":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Delivery Charge must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "RETURNCREDIT":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Return Credit must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "SELLPRICE":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Sell Price must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "REPLACEMENTCOST":
                            if (data != "")
                            {
                                bool isNumber = decimal.TryParse(data, out decimal tempNumber);
                                if (!isNumber)
                                {
                                    msg = "Replacement Cost must be a number or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "MIN.DAYS":
                            if (data != "")
                            {
                                isInt = int.TryParse(data, out tempInt);
                                if (!isInt || tempInt < 1)
                                {
                                    msg = "Min.Days must be an integer greater than zero. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;

                        case "RENTADJUSTMENT":
                            if (data != "")
                            {
                                bool found = ValidateSheetValue(data, adjustmentSheet, 1);
                                if (!found)
                                {
                                    msg = "Rent Adjustment not found in Rent Adjustment import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateCustomer(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundNumbers = new Dictionary<string, int>();

            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is required; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (foundEntries.ContainsKey(data))
                                {
                                    msg = "Name is a duplicate. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(data, 0);
                                }
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "NUMBER":
                            if (data != "")
                            {
                                if (foundNumbers.ContainsKey(data))
                                {
                                    msg = "Number is a duplicate. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundNumbers.Add(data, 0);
                                }
                            }
                            break;
                        case "ACCOUNTINGID":
                            break;
                        case "EMAIL":
                            break;
                        case "WEBSITE":
                            break;
                        case "BUSINESSPHONE":
                            break;
                        case "FAX#":
                            break;
                        case "BUSINESSSTREET1":
                            break;
                        case "BUSINESSSTREET2":
                            break;
                        case "BUSINESSCITY":
                            break;
                        case "BUSINESSSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Business State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSZIP/POSTALCODE":
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Business County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            break;
                        case "BUSINESSLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGSTREET1":
                            break;
                        case "BILLINGSTREET2":
                            break;
                        case "BILLINGCITY":
                            break;
                        case "BILLINGSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Billing State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Billing County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            break;
                        case "BILLINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGSTREET1":
                            break;
                        case "SHIPPINGSTREET2":
                            break;
                        case "SHIPPINGCITY":
                            break;
                        case "SHIPPINGSTATE":
                            if (data != "")
                            {

                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Shipping State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGZIP/POSTALCODE":
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            if (data != "")
                            {

                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Shipping County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            break;
                        case "SHIPPINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "SHIPPINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "SHIPPINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateCustomerContact(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string customerName = "";
            string first = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundUsernames = new Dictionary<string, int>();

            Excel.Worksheet customerSheet = FindSheet("IMPORT:CUSTOMER", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "CUSTOMERNAME":
                            if (!ValidateSheetValue(data, customerSheet, 1))
                            {
                                msg = "Customer Name not found in Customer import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            customerName = data;
                            break;
                        case "FIRST":
                            if (data == "")
                            {
                                msg = "First Name is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            first = data;
                            break;
                        case "LAST":
                            if (data == "")
                            {
                                msg = "Last Name is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (customerName != "" && first != "")
                            {
                                string nameKey = customerName + "|" + first + "|" + data;
                                if (foundEntries.ContainsKey(nameKey))
                                {
                                    msg = "[Customer Name, First, Last] can not be duplicated in the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(nameKey, 0);
                                }
                            }
                            break;
                        case "TITLE":
                            break;
                        case "EMAIL":
                            break;
                        case "PHONE":
                            break;
                        case "FAX":
                            break;
                        case "CELL":
                            break;
                        case "NOTES":
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "USERNAME":
                            string userKey = customerName + "|" + data;
                            if (data == "" || data.TrimEnd().Length < 5)
                            {
                                msg = "User Name is required and must have at least 5 characters. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundUsernames.ContainsKey(userKey))
                            {
                                msg = "[Customer Name, Username] can not be duplicated in the sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                foundUsernames.Add(userKey, 0);
                            }
                            break;
                        case "PASSWORD":
                            if (data == "")
                            {
                                msg = "Password is required; it must not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }

        }

        internal void ValidateOrder(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();

            Excel.Worksheet customerSheet = FindSheet("IMPORT:CUSTOMER", wb, sheetDic); //
            Excel.Worksheet customerContactSheet = FindSheet("IMPORT:CUSTOMERCONTACT", wb, sheetDic);


            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                string orderNumber = "";
                string workOrderNumber = "";
                string customerName = "";

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "CUSTOMERNAME":
                            if (data == "")
                            {
                                msg = "Customer Name is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (!ValidateSheetValue(data, customerSheet, 1))
                                {
                                    msg = "Customer Name not found in the Customer import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    customerName = data;
                                }
                            }
                            break;
                        case "ORDERNUMBER":
                            if (data == "")
                            {
                                msg = "Order Number is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                orderNumber = data;
                            }
                            break;
                        case "WORKORDERNUMBER":
                            if (orderNumber != "")
                            {
                                string orderKey = orderNumber + "|" + data;
                                if (foundEntries.ContainsKey(orderKey))
                                {
                                    msg = "[Order Number, Work Order Number] can not be repeated on the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(orderKey, 0);
                                    workOrderNumber = data;
                                }
                            }
                            break;
                        case "ORDERISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Order Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "ORDERDESCRIPTION":
                            break;
                        case "ORDERAPPROVERUSERNAME":
                            if (data != "" && customerName != "")   // Skip test if customerName is still ""
                            {
                                if (customerName != "")
                                {
                                    if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                    {
                                        msg = "Order Approver Username not found in Customer Contact import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                }
                                else
                                {
                                    msg = "Order Approver Userame can't be validated without a valid Customer Name. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ORDERREVIEWERUSERNAME":
                            if (data != "" && customerName != "")   // Skip test if customerName is still ""
                            {
                                if (customerName != "")
                                {
                                    if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                    {
                                        msg = "Order Reviewer Username not found in Customer Contact import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                }
                                else
                                {
                                    msg = "Order Reviewer Userame can't be validated without a valid Customer Name. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ORDERAMOUNT":
                            if (data != "")
                            {
                                bool isNumber = Decimal.TryParse(data, out decimal tempDecimal);
                                if (!isNumber)
                                {
                                    msg = "Order Amount must be a number or left blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "WORKORDERISACTIVE":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Order Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "WORKORDERDESCRIPTION":
                            break;
                        case "WORKORDERAPPROVERUSERNAME":
                            if (data != "" && customerName != "")   // Skip test if customerName is still ""
                            {
                                if (customerName != "")
                                {
                                    if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                    {
                                        msg = "Work Order Approver Username not found in Customer Contact import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                }
                                else
                                {
                                    msg = "Work Order Approver Username can't be validated without a valid Customer Name. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "WORKORDERREVIEWERUSERNAME":
                            if (data != "" && customerName != "")   // Skip test if customerName is still ""
                            {
                                if (customerName != "")
                                {
                                    if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                    {
                                        msg = "Work Order Reviewer Username not found in Customer Contact import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                }
                                else
                                {
                                    msg = "Work Order Reviewer Username can't be validated without a valid Customer Name. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "WORKORDERAMOUNT":
                            if (data != "")
                            {
                                bool isNumber = Decimal.TryParse(data, out decimal tempDecimal);
                                if (!isNumber)
                                {
                                    msg = "Work Order Amount must be a number or left blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateVendor(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundNumbers = new Dictionary<string, int>();

            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name is required; can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (foundEntries.ContainsKey(data))
                                {
                                    msg = "Name is a duplicate. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(data, 0);
                                }
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "NUMBER":
                            if (data != "")
                            {
                                if (foundNumbers.ContainsKey(data))
                                {
                                    msg = "Number is a duplicate. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundNumbers.Add(data, 0);
                                }
                            }
                            break;
                        case "ACCOUNTINGID":
                            break;
                        case "EMAIL":
                            break;
                        case "WEBSITE":
                            break;
                        case "BUSINESSPHONE":
                            break;
                        case "FAX#":
                            break;
                        case "BUSINESSSTREET1":
                            break;
                        case "BUSINESSSTREET2":
                            break;
                        case "BUSINESSCITY":
                            break;
                        case "BUSINESSSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Business State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSZIP/POSTALCODE":
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Business County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            break;
                        case "BUSINESSLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGSTREET1":
                            break;
                        case "BILLINGSTREET2":
                            break;
                        case "BILLINGCITY":
                            break;
                        case "BILLINGSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Billing State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Billing County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            break;
                        case "BILLINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGSTREET1":
                            break;
                        case "SHIPPINGSTREET2":
                            break;
                        case "SHIPPINGCITY":
                            break;
                        case "SHIPPINGSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Shipping State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGZIP/POSTALCODE":
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Shipping County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            break;
                        case "SHIPPINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "SHIPPINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "SHIPPINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateVendorContact(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string vendorName = "";
            string first = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundUsernames = new Dictionary<string, int>();

            Excel.Worksheet vendorSheet = FindSheet("IMPORT:VENDOR", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "VENDORNAME":
                            if (!ValidateSheetValue(data, vendorSheet, 1))
                            {
                                msg = "Vendor Name not found in Vendor import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            vendorName = data;
                            break;
                        case "FIRST":
                            if (data == "")
                            {
                                msg = "First Name is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            first = data;
                            break;
                        case "LAST":
                            if (data == "")
                            {
                                msg = "Last Name is required; it can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (vendorName != "" && first != "")
                            {
                                string nameKey = vendorName + "|" + first + "|" + data;
                                if (foundEntries.ContainsKey(nameKey))
                                {
                                    msg = "[Vendor Name, First, Last] can not be duplicated in the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(nameKey, 0);
                                }
                            }
                            break;
                        case "TITLE":
                            break;
                        case "EMAIL":
                            break;
                        case "PHONE":
                            break;
                        case "FAX":
                            break;
                        case "CELL":
                            break;
                        case "NOTES":
                            break;
                        case "ISACTIVE":
                            break;
                        case "USERNAME":
                            string userKey = vendorName + "|" + data;
                            if (data == "" || data.TrimEnd().Length < 5)
                            {
                                msg = "User Name is required and must have at least 5 characters. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else if (foundUsernames.ContainsKey(userKey))
                            {
                                msg = "[Vendor Name, Username] can not be duplicated in the sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                foundUsernames.Add(userKey, 0);
                            }
                            break;
                        case "PASSWORD":
                            if (data == "")
                            {
                                msg = "Password is required; it must not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateTaxRate(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            // Dictionary<string, int> foundEntriesx = new Dictionary<string, int>();
            HashSet<string> foundEntries = new HashSet<string>();

            Excel.Worksheet vendorSheet = FindSheet("IMPORT:VENDOR", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (foundEntries.Contains(data))
                            {
                                msg = "Name values  must be unique on the worksheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data);
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, true, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "REFID":
                            break;
                        case "TAXAGENCY":
                            break;
                        case "DESCRIPTION":
                            break;
                        case "RATE":
                            if (data != "")
                            {
                                bool isNumber = Double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "The Rate must be a valid number or blank.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateJobSite(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundNumbers = new Dictionary<string, int>();

            Excel.Worksheet customerSheet = FindSheet("IMPORT:CUSTOMER", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);
            Excel.Worksheet taxRateSheet = FindSheet("IMPORT:TAXRATE", wb, sheetDic);
            Excel.Worksheet rateProfileSheet = FindSheet("IMPORT:RATEPROFILE", wb, sheetDic);
            Excel.Worksheet orderSheet = FindSheet("IMPORT:ORDER", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                string customerName = "";
                string orderNumber = "";
                bool isBillable = false;
                bool advanceCycleCommonDate = false;
                bool trackScaffolds = false;
                DateTime firstInvoiceDate = DateTime.MinValue;
                int specificDay = 0;

                CycleType cycleType = CycleType.None;
                MonthlyBillingCycleType monthlyType = MonthlyBillingCycleType.None;

                BillingMethodType billingMethod = BillingMethodType.All;

                // Invoice Option is no longer optional - Always SingleOrder
                InvoiceOption invoiceOption = InvoiceOption.SingleOrder;

                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "PARENTLOCATION":
                            if (data == "")
                            {
                                msg = "Parent Location can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "CUSTOMERNAME":
                            if (data == "")
                            {
                                msg = "Customer Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (!ValidateSheetValue(data, customerSheet, 1))
                            {
                                msg = "Customer Name not found on Customer import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else customerName = data;
                            break;
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (foundEntries.ContainsKey(data))
                                {
                                    msg = "Name can not be repeated on the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(data, 0);
                                }
                            }
                            break;
                        case "NUMBER":
                            if (data != "")
                            {
                                if (foundNumbers.ContainsKey(data))
                                {
                                    msg = "Number can not be repeated on the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundNumbers.Add(data, 0);
                                }
                            }
                            break;
                        case "ACCOUNTINGID":
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "ISBILLABLE":
                            result = CellBool(data, false, true);
                            resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Billable must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                isBillable = (result == BoolResult.True);
                            }
                            break;
                        case "DESCRIPTION":
                            break;
                        case "TRACKSCAFFOLDS":
                            result = CellBool(data, false, true);
                            resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Track Scaffolds must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            else
                            {
                                trackScaffolds = result == BoolResult.True;
                            }
                            break;
                        case "TRACKACTIVITIES":
                            result = CellBool(data, false, true);
                            resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Track Activities must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "REQUESTS":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Requests must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST1LABEL":
                            break;
                        case "SCAFFOLDACTIVITYLIST1REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Activity List1 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST2LABEL":
                            break;
                        case "SCAFFOLDACTIVITYLIST2REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Activity List2 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST3LABEL":
                            break;
                        case "SCAFFOLDACTIVITYLIST3REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Activity List3 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDACTIVITYTEXTLABEL":
                            break;
                        case "SCAFFOLDACTIVITYTEXTREQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Activity Text Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST1LABEL":
                            break;
                        case "SCAFFOLDLIST1REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List1 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST2LABEL":
                            break;
                        case "SCAFFOLDLIST2REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List2 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST3LABEL":
                            break;
                        case "SCAFFOLDLIST3REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List3 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST4LABEL":
                            break;
                        case "SCAFFOLDLIST4REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List4 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST5LABEL":
                            break;
                        case "SCAFFOLDLIST5REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List5 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST6LABEL":
                            break;
                        case "SCAFFOLDLIST6REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List6 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDLIST7LABEL":
                            break;
                        case "SCAFFOLDLIST7REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold List7 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDDATE1LABEL":
                            break;
                        case "SCAFFOLDDATE1REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Date1 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDDATE2LABEL":
                            break;
                        case "SCAFFOLDDATE2REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Date2 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDTEXT1LABEL":
                            break;
                        case "SCAFFOLDTEXT1REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Text1 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDTEXT2LABEL":
                            break;
                        case "SCAFFOLDTEXT2REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Text2 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDTEXT3LABEL":
                            break;
                        case "SCAFFOLDTEXT3REQUIRED":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Text3 Required must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDYES/NO1LABEL":
                            break;
                        case "SCAFFOLDYES/NO2LABEL":
                            break;
                        case "BUSINESSSTREET1":
                            break;
                        case "BUSINESSSTREET2":
                            break;
                        case "BUSINESSCITY":
                            break;
                        case "BUSINESSSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Business State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSZIP/POSTALCODE":
                            break;
                        case "BUSINESSCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Business County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSCOUNTRY":
                            break;
                        case "BUSINESSLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Business Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BUSINESSELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGSTREET1":
                            break;
                        case "BILLINGSTREET2":
                            break;
                        case "BILLINGCITY":
                            break;
                        case "BILLINGSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Billing State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGZIP/POSTALCODE":
                            break;
                        case "BILLINGCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Billing County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGCOUNTRY":
                            break;
                        case "BILLINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Billing Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGSTREET1":
                            break;
                        case "SHIPPINGSTREET2":
                            break;
                        case "SHIPPINGCITY":
                            break;
                        case "SHIPPINGSTATE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STATES / PROVINCES", 1, norm1: true))
                                {
                                    msg = "Shipping State not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGZIP/POSTALCODE":
                            break;
                        case "SHIPPINGCOUNTY/PARISH":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "COUNTIES / PARISHES", 1, norm1: true))
                                {
                                    msg = "Business County/Parish not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGCOUNTRY":
                            break;
                        case "SHIPPINGLATITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Latitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGLONGITUDE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Longitude must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SHIPPINGELEVATION":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "Non-blank Shipping Elevation must be a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "JOBTEXT1":
                            break;
                        case "JOBTEXT2":
                            break;
                        case "JOBYESNO1":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Job Yes No 1 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "JOBYESNO2":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Job Yes No 2 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "JOBLIST1":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "JOBLIST1", 1, norm1: true))
                                {
                                    msg = "Job List 1 not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "JOBLIST2":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "JOBLIST2", 1, norm1: true))
                                {
                                    msg = "Job List 2 not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BILLINGMETHOD":
                            if (data == "" && isBillable)
                            {
                                msg = "Billing Method is required for billable job site. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (data != "")
                            {
                                switch (data)
                                {
                                    case "ARREARS":
                                        billingMethod = BillingMethodType.Arrears;
                                        break;
                                    case "FATA":
                                        billingMethod = BillingMethodType.FirstAdvancedThenArrears;
                                        break;
                                    case "ADVANCE":
                                        billingMethod = BillingMethodType.Advanced;
                                        break;
                                    default:
                                        msg = "Billing Method must be \"Arrears\", \"FATA\", or \"Advance\". ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                        billingMethod = BillingMethodType.All;
                                        break;
                                }
                            }
                            break;

                        case "FIRSTINVOICEDATE":
                            if (data == "" && billingMethod == BillingMethodType.Arrears)
                            {
                                msg = "First Invoice Date required for Arrears billing. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "First Invoice Date is not in valid date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    firstInvoiceDate = tempDate;
                                }
                            }
                            break;
                        case "FATAMINIMUMDAYSRENT":
                            if (billingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                bool isInteger = int.TryParse(data, out int tempInt);
                                if (!isInteger || tempInt <= 0)
                                {
                                    msg = "FATA Minimum Days Rent must be at least 1.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "FATA Minimum Days Rent is valid only for FATA Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "FATAALLOWOVERRIDE":
                            if (billingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                result = CellBool(data, true, true);
                                if (result == BoolResult.Invalid)
                                {
                                    msg = "FATA Allow Override must be Yes, No, True, False or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "FATA Allow Override is valid only for FATA Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "FATAINCLUDEADDITIONALCHARGES":
                            if (billingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                result = CellBool(data, true, true);
                                if (result == BoolResult.Invalid)
                                {
                                    msg = "FATA Include Additional Charges must be Yes, No, True, False or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "FATA Include Additional Charges is valid only for FATA Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ADVANCEISSUECREDITS":
                            if (billingMethod == BillingMethodType.Advanced)
                            {
                                result = CellBool(data, true, true);
                                if (result == BoolResult.Invalid)
                                {
                                    msg = "Advance Issue Credits must be Yes, No, True, False or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Advance Issue Credits is valid only for Advance Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ADVANCEMINIMUMDAYSRENT":
                            if (billingMethod == BillingMethodType.Advanced)
                            {
                                if (data != "")
                                {
                                    bool isInt = int.TryParse(data, out int tempInt);
                                    if (!isInt)
                                    {
                                        msg = "Advance Minimum Days Rent must be an integer or blank.  ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                    }
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Advance Minimum Days Rent is valid only for Advance Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ADVABCEINVOICECYCLEDAYS":
                            if (billingMethod == BillingMethodType.Advanced)
                            {
                                if (data != "")
                                {
                                    bool isInt = int.TryParse(data, out int tempInt);
                                    if (!isInt || tempInt <= 0)
                                    {
                                        msg = "Advance Invoice Cycle Days must be an integer and at least 1.  ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                    }
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Advance Invoice Cycle Days is valid only for Advance Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ADVANCECYCLECOMMONDATE":
                            if (billingMethod == BillingMethodType.Advanced)
                            {
                                result = CellBool(data, true, true);
                                if (result == BoolResult.Invalid)
                                {
                                    msg = "Advance Cycle Common Date must be Yes, No, True, False or blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                                else
                                {
                                    advanceCycleCommonDate = (result == BoolResult.True);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Advance Cycle Common Date  is valid only for Advance Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ADVANCECYCLESTARTDATE":
                            if (billingMethod == BillingMethodType.Advanced)
                            {
                                if (data != "")
                                {
                                    bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                    if (!isDate)
                                    {
                                        msg = "Advance Cycle Start Date must be a valid date.  ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                    }
                                }
                                else if (advanceCycleCommonDate)
                                {
                                    msg = "Advance Cycle Start Date  is required when Advance Cycle Common Date is true.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Advance Cycle Start Date is valid only for Advance Billing Method.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "DAILY/MONTHLY":
                            if (billingMethod == BillingMethodType.Arrears || billingMethod == BillingMethodType.FirstAdvancedThenArrears)
                            {
                                switch (data)
                                {
                                    case "DAILY":
                                        cycleType = CycleType.Daily;
                                        break;
                                    case "MONTHLY":
                                        cycleType = CycleType.Monthly;
                                        break;
                                    default:
                                        msg = "Daily / Monthly must be \"Daily\" or \"Monthly\" for Arrears or FATA billing. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                        break;
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Daily / Monthly  is valid only for Arrears or FATA billing.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "INVOICEEVERYNUMBEROFDAYS":
                            if (cycleType == CycleType.Daily)
                            {
                                bool isInt = int.TryParse(data, out int tempInt);
                                if (!isInt || tempInt <= 0)
                                {
                                    msg = "Invoice Every Number Of Days must be an integer greater than zero for Daily billing.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }

                            }
                            else if (data != "")
                            {
                                msg = "Invoice Every Number Of Days is valid only for Daily billing.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "INVOICEMONTHLYCYCLETYPE":
                            if (cycleType == CycleType.Monthly)
                            {

                                switch (data)
                                {
                                    case "LAST DAY":
                                        monthlyType = MonthlyBillingCycleType.LastDay;
                                        break;
                                    case "FIRST DAY":
                                        monthlyType = MonthlyBillingCycleType.FirstDay;
                                        break;
                                    case "LAST SUNDAY":
                                        monthlyType = MonthlyBillingCycleType.LastSunday;
                                        break;
                                    case "SPECIFIC DAY":
                                        monthlyType = MonthlyBillingCycleType.SpecificDay;
                                        break;
                                    default:
                                        msg = "Invoice Monthly Cycle Type must be \"Last Day\", \"First Day\", \"Last Sunday\", or \"Specific Day\" for Monthly billing.  ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                        break;
                                }
                            }
                            else if (data != "")
                            {
                                msg = "Invoice Monthly Cycle is valid only for Monthly billing.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "INVOICEMONTHLYDAYOFMONTH":
                            if (data != "" && (cycleType != CycleType.Monthly || monthlyType != MonthlyBillingCycleType.SpecificDay))
                            {
                                msg = "Invoice Monthly Day of Month is valid only for monthly billing with Monthly Type of Specific Day.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (data == "" && cycleType == CycleType.Monthly && monthlyType == MonthlyBillingCycleType.SpecificDay)
                            {
                                msg = "Invoice Monthly Day of Month is required for Monthly billing on a Specific Day.  ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (data != "")
                            {
                                bool isInt = int.TryParse(data, out int tempInt);
                                if (!isInt || tempInt <= 0 || tempInt >= 31)
                                {
                                    msg = "Invoice Monthly Day of Month must be an integer between 1 and 31.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    specificDay = tempInt;
                                }
                            }

                            if (cycleType == CycleType.Monthly && billingMethod == BillingMethodType.Arrears)
                            {
                                bool goodDate = CalculatedDateInMonth(firstInvoiceDate, monthlyType, specificDay, out DateTime calcDate);
                                if (!goodDate || calcDate != firstInvoiceDate)
                                {
                                    msg = "First Invoice Date for Arrears billing is inconsistent with Monthly Cycle Type.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "RATEPROFILE":
                            if (isBillable)
                            {
                                if (data == "")
                                {
                                    msg = "Rate Profile is required if job site is billable.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else if (!ValidateSheetValue(data, rateProfileSheet, 1))
                                {
                                    msg = "Rate Profile is not found on Rate Profile import sheet.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PURCHASEORDER":
                            if (data == "" && trackScaffolds && isBillable)
                            {
                                msg = "A billable, scaffold tracking job site must have a purchase order. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (data != "")
                            {
                                orderNumber = data;
                            }
                            break;
                        case "WORKORDER":
                            if (data != "" && orderNumber == "")
                            {
                                msg = "If Purchase Order is blank, then Work Order must also be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (orderNumber != "")
                            {
                                if (data == "")
                                {
                                    if (!ValidateSheetValue(orderNumber, orderSheet, 2, customerName, 1, "", 3))
                                    {
                                        msg = "Purchase Order was not found on Order import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                    else if (!ValidateSheetValue(data, orderSheet, 3, customerName, 1, orderNumber, 2))
                                    {
                                        msg = "[Purchase Order, Work Order] was not found on Order import sheet. ";
                                        WriteCell(row, resultCol, ws, msg, append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                    }
                                }
                            }
                            break;
                        case "DEFAULTNEWSHIPMENTSBILLABLE":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Default New Shipments Billable must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "HIDEZEROCHARGES":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Hide Zero Charges must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "TAXRATE1":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, taxRateSheet, 1))
                                {
                                    msg = "Tax Rate 1 not found on Tax Rate import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "TAXRATE2":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, taxRateSheet, 1))
                                {
                                    msg = "Tax Rate 2 not found on Tax Rate import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                                }
                            }
                            break;
                        case "TAXCONSUMABLESANDPRODUCTSALES":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Tax Consumables and Product Sales must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "TAXRENT":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Tax Rent must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "TAXSERVICETICKETDAMAGE":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Tax Service Ticket Damage must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "TAXDELIVERY&RETURN":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Tax Delivery and Return must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateRequest(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            string jobSiteName = "";
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            Dictionary<string, int> foundNumbers = new Dictionary<string, int>();

            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet customerContactSheet = FindSheet("IMPORT:CUSTOMERCONTACT", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "REQUESTNUMBER":
                            if (data == "")
                            {
                                msg = "Request Number can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (foundEntries.ContainsKey(data))
                                {
                                    msg = "Request Number can not be repeated on the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(data, 0);
                                }
                            }
                            break;
                        case "NOTES":
                            if (data == "")
                            {
                                msg = "Notes can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ORDER":
                            break;
                        case "STATUS":
                            break;
                        case "ASSIGNEDTOUSERNAME":
                            break;
                        case "REQUESTDATE":
                            bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                            if (!isDate)
                            {
                                msg = "Request Date is required a must be in date format. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "DATENEEDED":
                            isDate = DateTime.TryParse(data, out tempDate);
                            if (!isDate)
                            {
                                msg = "Date Needed is required a must be in date format. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "REQUESTLIST1":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "REQUESTLIST1", 1, norm1: true))
                                {
                                    msg = "Request List 1 not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "REQUESTLIST2":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "REQUESTLIST2", 1, norm1: true))
                                {
                                    msg = "Request List 2 not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "REQUESTLIST3":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "REQUESTLIST3", 1, norm1: true))
                                {
                                    msg = "Request List 3 not found in Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "REQUESTTEXT1":
                            break;
                        case "REQUESTTEXT2":
                            break;
                        case "REQUESTYESNO1":
                            BoolResult result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Request YesNo 1 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "REQUESTYESNO2":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Request YesNo 2 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateScaffoldJobList(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string listName = "";
            string jobSiteName = "";

            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Dictionary<string, int> foundEntries = new Dictionary<string, int>();
            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "LIST":
                            if (data == "")
                            {
                                msg = "List value can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                listName = NormStr(data);
                                if (!ValidScaffoldJobListNames.Contains(listName))
                                {
                                    msg = "List value is not a valid Scaffold Job List value. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }

                            break;
                        case "VALUE":
                            if (data == "")
                            {
                                msg = "Value column can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                string valueKey = jobSiteName + "|" + listName + "|" + data;
                                if (foundEntries.ContainsKey(valueKey))
                                {
                                    msg = "[Job Site Name, List, Value] can not be duplicated in the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(valueKey, 0);
                                }
                            }
                            break;
                        case "DESCRIPTION":
                            break;
                    }
                }
            }
        }

        internal void ValidateScaffoldTag(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
          Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();


            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet scaffoldJobListSheet = FindSheet("IMPORT:SCAFFOLDJOBLIST", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);
            Excel.Worksheet orderSheet = FindSheet("IMPORT:ORDER", wb, sheetDic);
            Excel.Worksheet customerContactSheet = FindSheet("IMPORT:CUSTOMERCONTACT", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                string jobSiteName = "";
                string customerName = "";
                string orderNumber = "";

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "PURCHASEORDER":
                            if (data != "")
                            {
                                orderNumber = data;
                            }
                            break;
                        case "WORKORDER":
                            if (orderNumber == "")
                            {
                                if (data != "")
                                {
                                    msg = "If Work Order is not blank, then Purchase Order must also be non-blank. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            else
                            {
                                customerName = CustNameFromJobSiteName(jobSiteName, jobsiteSheet);
                                if (customerName == "")
                                {
                                    msg = "Customer Name not found in the Job Site import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "STATUS":
                            if (data == "")
                            {
                                // bug 15847 dropped requirement that we validate the status value
                                msg = "Status is required; it can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "TAG":
                            if (data == "")
                            {
                                msg = "Tag can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                if (foundEntries.ContainsKey(data))
                                {
                                    msg = "Tag can not be repeated on the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(data, 0);
                                }
                            }
                            break;
                        case "PRIORITY":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "PRIORITY", 1))
                                {
                                    msg = "Priority not found on Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PROJECT":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "PROJECT", 1))
                                {
                                    msg = "Project not found on Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "STEP":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, globalListsSheet, 2, "STEP", 1))
                                {
                                    msg = "Step not found on Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "REQUESTED BY":
                            if (data != "" && data != "CUSTOMER" && data != "INTERNAL")
                            {
                                msg = "Requested By must be Customer, Internal or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;

                        case "REPRESENTATIVE":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                {
                                    msg = "Representative not found on the Customer Contact import sheet for the job site's customer. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "REQUESTOR":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, customerContactSheet, 11, customerName, 1))
                                {
                                    msg = "Requestor not found on the Customer Contact import sheet for the job site's customer. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;

                        case "FOREMAN":
                            break;
                        case "BUILDER":
                            break;
                        case "DISMANTLER":
                            break;
                        case "INSPECTOR":
                            break;

                        case "PLANNEDLOADDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Planned Load Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PLANNEDBUILDDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Planned Build Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PLANNEDDISMANTLEDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Planned Dismantle Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ACTUALLOADDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Actual Load Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ACTUALBUILDDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Actual Build Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ACTUALDISMANTLEDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Actual Dismantle Date is not a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "NOTES":
                            break;
                        case "LOCATIONNOTES":
                            break;
                        case "LENGTH":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "A non-blank Length must be a number.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "WIDTH":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "A non-blank Width must be a number.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "HEIGHT":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "A non-blank Height must be a number.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "NOOFLEGS":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "A non-blank No of Legs must be a number.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "NOOFDECKS":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempNumber);
                                if (!isNumber)
                                {
                                    msg = "A non-blank No of Decks must be a number.  ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "BASEELEVATION":
                            break;
                        case "SCAFFOLDLIST1":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST1", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List1 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST2":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST2", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List2 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST3":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST3", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List3 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST4":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST4", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List4 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST5":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST5", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List5 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST6":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST6", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List6 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDLIST7":
                            if (data != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLDLIST7", 2, jobSiteName, 1, norm1: true))
                                {
                                    msg = "Scaffold List7 value not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }

                            }
                            break;
                        case "SCAFFOLDDATE1":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Scaffold Date 1 must be a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDDATE2":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non-blank Scaffold Date 2 must be a valid date. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDTEXT1":
                            break;
                        case "SCAFFOLDTEXT2":
                            break;
                        case "SCAFFOLDTEXT3":
                            break;
                        case "SCAFFOLDYES/NO1":
                            BoolResult result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Yes/No 1 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;

                        case "SCAFFOLDYES/NO2":
                            result = CellBool(data, true, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Scaffold Yes/No 1 must be Yes, No, True, False or blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateScaffoldTagActivity(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            Dictionary<string, int> foundEntries = new Dictionary<string, int>();

            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet scaffoldJobListSheet = FindSheet("IMPORT:SCAFFOLDJOBLIST", wb, sheetDic);
            Excel.Worksheet scaffoldTagSheet = FindSheet("IMPORT:SCAFFOLDTAG", wb, sheetDic);
            Excel.Worksheet globalListsSheet = FindSheet("IMPORT:GLOBALLISTS", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                string jobSiteName = "";
                string tag = "";
                string activityType = "";
                bool foundDate = false;

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "TAG":
                            if (data == "")
                            {
                                msg = "Tag can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (jobSiteName != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldTagSheet, 5, jobSiteName, 1))
                                {
                                    msg = "[Jobsite Name, Tag] not found on Scaffold Tag import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    tag = data;
                                }
                            }
                            break;
                        case "SCAFFOLDTAGACTIVITYTYPE":
                            if (data == "")
                            {
                                msg = "Scaffold Tag Activity Type can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (jobSiteName != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, jobSiteName, 1, "SCAFFOLDTAGACTIVITYTYPE", 2, norm2: true))
                                {
                                    msg = "Scaffold Tag Activity Type not found on Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    activityType = data;
                                }

                            }
                            break;
                        case "SUMMARY":
                            if (data == "")
                            {
                                msg = "Summary is required and can't be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                string entryKey = (jobSiteName + "|" + tag + "|" + activityType + "|" + data).ToUpper();
                                if (foundEntries.ContainsKey(entryKey))
                                {
                                    msg = "[Job Site Name, Tag, Scaffold Tag Activity Type, Summary] can't be duplicated in the sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(entryKey, 0);
                                }
                            }
                            break;
                        case "DETAIL":
                            break;
                        case "SHIFT":
                            if (data != "")
                            {
                                // Accept "NIGHT" AND "DAY" without checking the global Lists Sheet
                                if (data != "DAY" && data != "NIGHT" && !ValidateSheetValue(data, globalListsSheet, 2, "SHIFTS", 1))
                                {
                                    msg = "Can't find Shift in the Global Lists import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "REQUESTDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non blank Request Date must be in date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundDate = true;
                                }
                            }
                            break;
                        case "PLANNEDDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non blank Planned Date must be in date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundDate = true;
                                }
                            }
                            break;
                        case "ACTUALDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non blank Actual Date must be in date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundDate = true;
                                }
                            }
                            break;
                        case "PLANNEDLOADDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non blank Planned Load Date must be in date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundDate = true;
                                }
                            }
                            break;
                        case "ACTUALLOADDATE":
                            if (data != "")
                            {
                                bool isDate = DateTime.TryParse(data, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Non blank Actual Load Date must be in date format. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundDate = true;
                                }
                            }

                            if (!foundDate)
                            {
                                msg = "At least one date must be provided. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST1":
                            if (data != "" && jobSiteName != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLD ACTIVITY LIST1", 2, jobSiteName, 1))
                                {
                                    msg = "Can't find Scaffold Activity List1 in the Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST2":
                            if (data != "" && jobSiteName != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLD ACTIVITY LIST2", 2, jobSiteName, 1))
                                {
                                    msg = "Can't find Scaffold Activity List2 in the Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYLIST3":
                            if (data != "" && jobSiteName != "")
                            {
                                if (!ValidateSheetValue(data, scaffoldJobListSheet, 3, "SCAFFOLD ACTIVITY LIST3", 2, jobSiteName, 1))
                                {
                                    msg = "Can't find Scaffold Activity List3 in the Scaffold Job List import sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SCAFFOLDACTIVITYTEXT":
                            break;
                        case "PLANNEDHOURS":
                            if (data != "" && jobSiteName != "")
                            {
                                bool isDouble = double.TryParse(data, out double tempDouble);
                                if (!isDouble)
                                {
                                    msg = "Planned Hours is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PLANNEDCOST":
                            if (data != "" && jobSiteName != "")
                            {
                                bool isDecimal = decimal.TryParse(data, out decimal tempDecimal);
                                if (!isDecimal)
                                {
                                    msg = "Planned Cost is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ACTUALHOURS":
                            if (data != "" && jobSiteName != "")
                            {
                                bool isDouble = double.TryParse(data, out double tempDouble);
                                if (!isDouble)
                                {
                                    msg = "Actual Hours is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "ACTUALCOST":
                            if (data != "" && jobSiteName != "")
                            {
                                bool isDecimal = decimal.TryParse(data, out decimal tempDecimal);
                                if (!isDecimal)
                                {
                                    msg = "Actual Cost is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateUnitOfMeasure(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            HashSet<string> foundEntries = new HashSet<string>();

            //Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);


            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (foundEntries.Contains(data))
                            {

                                msg = "Duplicate Name not allowed. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(data);
                            }
                            break;
                        case "DESCRIPTION":
                            if (data == "")
                            {
                                msg = "Description is required. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateUnitHourRateProfile(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            HashSet<string> foundEntries = new HashSet<string>();

            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet unitOfMeasureSheet = FindSheet("IMPORT:UNITOFMEASURE", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                string jobSiteName = "";

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {

                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }

                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on the Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "UNITHOURRATEPROFILE":
                            if (data == "")
                            {
                                msg = "Unit Hour Rate Profile can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else if (foundEntries.Contains(jobSiteName + "|" + data))
                            {
                                msg = "Duplicate Unit Hour Rate Profile for a Job Site is not allowed. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                foundEntries.Add(jobSiteName + "|" + data);
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                        case "UNITOFMEASURE":
                            if (data == "")
                            {
                                msg = "Unit of Measure Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }

                            else if (!ValidateSheetValue(data, unitOfMeasureSheet, 1))
                            {
                                msg = "Unit of Measure Name not found on the Unit of Measure import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;

                        case "UNITFACTOR":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double temp);
                                if (!isNumber)
                                {
                                    msg = "Unit Factor is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "SPLIT":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double temp);
                                if (!isNumber)
                                {
                                    msg = "Split is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                        case "PRICE":
                            if (data != "")
                            {
                                bool isDecimal = decimal.TryParse(data, out decimal temp);
                                if (!isDecimal)
                                {
                                    msg = "Price is not a decimal number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateMultiplier(Excel.Workbook wb, Excel.Worksheet ws, int maxRows, int resultCol, ref int validationRow,
            Excel.Worksheet validationSheet, ImportConfig importConfig, Dictionary<string, string> sheetDic)
        {
            string msg = "";

            HashSet<string> foundEntries = new HashSet<string>();

            Excel.Worksheet jobsiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet unitHourRateSheet = FindSheet("IMPORT:UNITHOURRATE", wb, sheetDic);

            for (int row = HeaderRow + 1; row <= maxRows; row++)
            {
                // Skip blank rows
                if (EmptyRow(ws, row, resultCol))
                {
                    continue;
                }
                string jobSiteName = "";

                foreach (var column in importConfig.ImportColumns)
                {
                    string data = CellData(ws, row, column.Column).ToUpper();
                    switch (NormStr(column.ColumnHeading))
                    {
                        case "JOBSITENAME":
                            if (data == "")
                            {
                                msg = "Job Site Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }

                            else if (!ValidateSheetValue(data, jobsiteSheet, 3))
                            {
                                msg = "Job Site Name not found on the Job Site import sheet. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            else
                            {
                                jobSiteName = data;
                            }
                            break;
                        case "NAME":
                            if (data == "")
                            {
                                msg = "Name can not be blank. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }

                            else
                            {
                                string lookupKey = jobSiteName + "|" + data;

                                if (foundEntries.Contains(lookupKey))
                                {
                                    msg = "[Job Site Name,  Name] can not be duplicated on this sheet. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                                else
                                {
                                    foundEntries.Add(lookupKey);
                                }
                            }
                            break;
                        case "VALUE":
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double temp);
                                if (!isNumber)
                                {
                                    msg = "Value is not a number. ";
                                    WriteCell(row, resultCol, ws, msg, append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                                }
                            }
                            else
                            {
                                msg = "Value is required. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1} ", row, msg), append: true);
                            }
                            break;
                        case "ISACTIVE":
                            BoolResult result = CellBool(data, false, true);
                            bool resultTrue = result == BoolResult.True;
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Active must be Yes, No, True, or False. ";
                                WriteCell(row, resultCol, ws, msg, append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, {1}", row, msg), append: true);
                            }
                            break;
                    }
                }
            }
        }

        internal void ValidateScaffoldTagBalance(Excel.Workbook wb, ref int validationRow, Excel.Worksheet validationSheet,
            Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string data = "";
            bool isBillable = false;

            Excel.Worksheet catalogSheet = FindSheet("IMPORT:PRODUCTCATALOG", wb, sheetDic);
            Excel.Worksheet scaffoldTagSheet = FindSheet("IMPORT:SCAFFOLDTAG", wb, sheetDic);
            Excel.Worksheet scaffoldTagActivitySheet = FindSheet("IMPORT:SCAFFOLDTAGACTIVITY", wb, sheetDic);
            Excel.Worksheet importBalanceSheet = FindSheet("IMPORT:SCAFFOLDTAGBALANCES", wb, sheetDic);

            WriteCell(++validationRow, 1, validationSheet, importBalanceSheet.Name);
            var activeRange = (object[,])importBalanceSheet.UsedRange.Value;
            if (activeRange == null)
            {

                WriteCell(validationRow++, 2, validationSheet, "Error: nothing to import");
                return;
            }
            int maxBalanceRow = activeRange.GetLength(0);
            int maxBalanceColumn = activeRange.GetLength(1);
            int resultRow = 0;

            for (int row = maxBalanceRow; row >= 14; row--)
            {
                data = CellData(importBalanceSheet, row, 1);
                if (data == Results)
                {
                    maxBalanceRow = row - 1;
                    resultRow = row;
                    break;
                }
                else if (data != "")
                {
                    maxBalanceRow = row;
                    resultRow = row + 1;
                    WriteCell(resultRow, 1, importBalanceSheet, Results);
                    break;
                }
            }
            if (resultRow == 0)
            {
                resultRow = 16;
            }
            else
            {
                WriteCell(resultRow, 2, importBalanceSheet, "");

                if (importBalanceSheet != null)
                {
                    if (CellData(importBalanceSheet, 7, 1) != "" ||
                        CellData(importBalanceSheet, 8, 1) != "" ||
                        CellData(importBalanceSheet, 9, 1) != "" ||
                        CellData(importBalanceSheet, 10, 1) != "" ||
                        CellData(importBalanceSheet, 11, 1) != "" ||
                        CellData(importBalanceSheet, 7, 2) != "Rent Start:" ||
                        CellData(importBalanceSheet, 8, 2) != "Job Site:" ||
                        CellData(importBalanceSheet, 9, 2) != "Tag:" ||
                        CellData(importBalanceSheet, 10, 2) != "Activity:" ||
                        CellData(importBalanceSheet, 11, 2) != "Summary:" ||
                        CellData(importBalanceSheet, 12, 2) != "Is Billable:" ||
                        CellData(importBalanceSheet, 13, 1) != "PartNumber" ||
                        CellData(importBalanceSheet, 13, 2) != "Description")
                    {
                        msg = "The Scaffold Tag Balance Sheet doesn't conform to the template for this import. Check the Row Headings. ";
                        WriteCell(resultRow, 2, importBalanceSheet, msg, append: true);
                        WriteCell(validationRow++, 2, validationSheet, msg);
                    }

                    for (int row = 14; row <= maxBalanceRow; row++)
                    {
                        data = CellData(importBalanceSheet, row, 1);
                        if (data == "")
                        {
                            msg = "PartNumber can not be blank. ";
                            WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                        }
                        else
                        {
                            if (!ValidateSheetValue(data, catalogSheet, 1))
                            {
                                msg = "Part Number not found on Product Catalog import sheet. ";
                                WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                            }
                        }
                    }

                    for (int column = 3; column <= maxBalanceColumn; column++)
                    {
                        if (!NoImportQuantities(importBalanceSheet, 14, resultRow, column))
                        {
                            WriteCell(resultRow, column, importBalanceSheet, "");
                            string startDate = CellData(importBalanceSheet, 7, column);
                            if (startDate == "")
                            {
                                msg = "Rent Start date is required and not found. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 7, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 7, ColumnIndexToLetter(column), msg), append: true);
                            }
                            else
                            {
                                bool isDate = DateTime.TryParse(startDate, out DateTime tempDate);
                                if (!isDate)
                                {
                                    msg = "Rent Start date is not a valid date. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 7, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 7, ColumnIndexToLetter(column), msg), append: true);
                                }
                            }

                            string jobSite = CellData(importBalanceSheet, 8, column);
                            if (jobSite == "")
                            {
                                msg = "Job Site is required and not found. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                            }

                            string tag = CellData(importBalanceSheet, 9, column);
                            if (tag == "")
                            {
                                msg = "Tag is required and not found. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 9, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 9, ColumnIndexToLetter(column), msg), append: true);
                            }
                            else
                            {
                                if (!ValidateSheetValue(tag, scaffoldTagSheet, 5, jobSite, 1))
                                {
                                    msg = "[Job Site, Tag] not found on the Scaffold Tag import sheet. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 9, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 9, ColumnIndexToLetter(column), msg), append: true);
                                }
                            }

                            string activity = CellData(importBalanceSheet, 10, column);

                            string summary = CellData(importBalanceSheet, 11, column);
                            if (activity != "" && summary == "")
                            {
                                msg = "If Activity is not blank, then Summary must be entered. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 10, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 10, ColumnIndexToLetter(column), msg), append: true);
                            }
                            else if (activity != "")
                            {
                                if (!ValidateSheetValue(activity, scaffoldTagActivitySheet, 3, jobSite, 1, tag, 2, summary, 4))
                                {
                                    msg = "[Job Site, Tag, Scaffold Tag Activity Type, Summary] not found on the Scaffold Tag Activity import sheet. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 10, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 10, ColumnIndexToLetter(column), msg), append: true);
                                }
                            }

                            data = CellData(importBalanceSheet, 12, column);
                            if (data != "")
                            {
                                BoolResult result = CellBool(data, false, true);
                                if (result == BoolResult.Invalid)
                                {
                                    msg = "Is Billable must be True, False, Yes, or No. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                                }
                                else
                                {
                                    isBillable = result == BoolResult.True;
                                }
                            }

                            data = CellData(importBalanceSheet, 13, column);
                            if (data != "")
                            {
                                msg = "Row 13 should be blank. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 15, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 15, ColumnIndexToLetter(column), msg), append: true);
                            }

                            for (int row = 14; row <= maxBalanceRow; row++)
                            {
                                data = CellData(importBalanceSheet, row, column);
                                if (data != "")
                                {
                                    bool isNumber = double.TryParse(data, out double tempDouble);
                                    if (!isNumber)
                                    {
                                        msg = "A non-blank quantity must be a valid number. ";
                                        WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                        WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", row, ColumnIndexToLetter(column), msg), append: true);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal void ValidateJobSiteBalance(Excel.Workbook wb, ref int validationRow, Excel.Worksheet validationSheet,
            Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string data = "";
            bool isBillable = false;


            Excel.Worksheet catalogSheet = FindSheet("IMPORT:PRODUCTCATALOG", wb, sheetDic);
            Excel.Worksheet jobSiteSheet = FindSheet("IMPORT:JOBSITE", wb, sheetDic);
            Excel.Worksheet importBalanceSheet = FindSheet("IMPORT:JOBSITEBALANCES", wb, sheetDic);

            ++validationRow;

            WriteCell(++validationRow, 1, validationSheet, importBalanceSheet.Name);
            var activeRange = (object[,])importBalanceSheet.UsedRange.Value;
            if (activeRange == null)
            {

                WriteCell(validationRow++, 2, validationSheet, "Error: nothing to import");
                return;
            }
            int maxBalanceRow = activeRange.GetLength(0);
            int maxBalanceColumn = activeRange.GetLength(1);
            int resultRow = 0;

            for (int row = maxBalanceRow; row >= 10; row--)
            {
                data = CellData(importBalanceSheet, row, 1);
                if (data == Results)
                {
                    maxBalanceRow = row - 1;
                    resultRow = row;
                    break;
                }
                else if (data != "")
                {
                    maxBalanceRow = row;
                    resultRow = row + 1;
                    WriteCell(resultRow, 1, importBalanceSheet, Results);
                    break;
                }
            }
            if (resultRow == 0)
            {
                resultRow = 11;
            }
            else
            {
                WriteCell(resultRow, 2, importBalanceSheet, "");

                if (importBalanceSheet != null)
                {
                    WriteCell(validationRow, 1, validationSheet, importBalanceSheet.Name);

                    if (CellData(importBalanceSheet, 7, 1) != "" ||
                        CellData(importBalanceSheet, 8, 1) != "" ||
                        CellData(importBalanceSheet, 9, 1) != "" ||
                        CellData(importBalanceSheet, 7, 2) != "Rent Start:" ||
                        CellData(importBalanceSheet, 8, 2) != "Job Site:" ||
                        CellData(importBalanceSheet, 9, 2) != "Is Billable:" ||
                        CellData(importBalanceSheet, 10, 1) != "PartNumber" ||
                        CellData(importBalanceSheet, 10, 2) != "Description")
                    {
                        msg = "The Job Site Balance Sheet doesn't conform to the template for this import. Check the Row Headings. ";
                        WriteCell(resultRow, 2, importBalanceSheet, msg, append: true);
                        WriteCell(validationRow++, 2, validationSheet, msg);
                    }

                    for (int row = 11; row <= maxBalanceRow; row++)
                    {
                        data = CellData(importBalanceSheet, row, 1);
                        if (data == "" && !EmptyRow(importBalanceSheet, row, maxBalanceColumn - 1))
                        {
                            msg = "PartNumber can not be blank. ";
                            WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                        }
                        else if (data != "")
                        {
                            if (!ValidateSheetValue(data, catalogSheet, 1))
                            {
                                msg = "Part Number not found on Product Catalog import sheet. ";
                                WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                            }
                        }
                    }

                    int newMaxCol = maxBalanceColumn;
                    for (int col = maxBalanceColumn; col >= 3; col--)
                    {
                        if (!EmptyCol(importBalanceSheet, col))
                        {
                            break;
                        }
                        else
                        {
                            newMaxCol -= 1;
                        }
                    }
                    maxBalanceColumn = newMaxCol;


                    for (int column = 3; column <= maxBalanceColumn; column++)
                    {
                        WriteCell(resultRow, column, importBalanceSheet, "");
                        string startDate = CellData(importBalanceSheet, 7, column);
                        if (startDate == "")
                        {
                            msg = "Rent Start date is required and not found. ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 7, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 7, ColumnIndexToLetter(column), msg), append: true);
                        }
                        else
                        {
                            bool isDate = DateTime.TryParse(startDate, out DateTime tempDate);
                            if (!isDate)
                            {
                                msg = "Rent Start date is not a valid date. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 7, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 7, ColumnIndexToLetter(column), msg), append: true);
                            }
                        }

                        string jobSite = CellData(importBalanceSheet, 8, column);
                        if (jobSite == "")
                        {
                            msg = "Job Site is required and not found. ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                        }
                        else
                        {
                            string customerName = CustNameFromJobSiteName(jobSite, jobSiteSheet);
                            if (customerName == "" || !ValidateSheetValue(jobSite, jobSiteSheet, 3, customerName, 2))
                            {
                                msg = "[Customer Name, Job Site] not found on Job Site import sheet. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                            }
                        }

                        data = CellData(importBalanceSheet, 9, column);
                        if (data != "")
                        {
                            BoolResult result = CellBool(data, false, true);
                            if (result == BoolResult.Invalid)
                            {
                                msg = "Is Billable must be True, False, Yes,  or No. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                            }
                            else
                            {
                                isBillable = result == BoolResult.True;
                            }
                        }

                        data = CellData(importBalanceSheet, 10, column);
                        if (data != "")
                        {
                            msg = "Row 10 should be blank. ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 10, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 10, ColumnIndexToLetter(column), msg), append: true);
                        }

                        for (int row = 11; row <= maxBalanceRow; row++)
                        {
                            data = CellData(importBalanceSheet, row, column);
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "A non-blank quantity must be a valid number. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", row, ColumnIndexToLetter(column), msg), append: true);
                                }
                            }
                        }
                    }
                }
            }
        }

        internal void ValidateBranchOrLaydownBalance(Excel.Workbook wb, ref int validationRow, Excel.Worksheet validationSheet,
            Dictionary<string, string> sheetDic)
        {
            string msg = "";
            string data = "";
            Excel.Worksheet catalogSheet = FindSheet("IMPORT:PRODUCTCATALOG", wb, sheetDic);
            Excel.Worksheet importBalanceSheet = FindSheet("IMPORT:BRANCHORLAYDOWNBALANCES", wb, sheetDic);
            HashSet<string> seenNames = new HashSet<string>();

            validationRow++;
            WriteCell(++validationRow, 1, validationSheet, importBalanceSheet.Name);
            var activeRange = (object[,])importBalanceSheet.UsedRange.Value;
            if (activeRange == null)
            {

                WriteCell(validationRow++, 2, validationSheet, "Error: nothing to import");
                return;
            }
            int maxBalanceRow = activeRange.GetLength(0);
            int maxBalanceColumn = activeRange.GetLength(1);
            for (int col = maxBalanceColumn; col > 2; col--)
            {
                for (int row = 1; row <= maxBalanceRow; row++)
                {
                    if (CellData(importBalanceSheet, row, col) != "")
                    {
                        maxBalanceColumn = col;
                        break;
                    }
                }
            }
            int resultRow = 0;

            for (int row = maxBalanceRow; row >= 11; row--)
            {
                data = CellData(importBalanceSheet, row, 1);
                if (data == Results)
                {
                    maxBalanceRow = row - 1;
                    resultRow = row;
                    break;
                }
                else if (data != "")
                {
                    maxBalanceRow = row;
                    resultRow = row + 1;
                    WriteCell(resultRow, 1, importBalanceSheet, Results);
                    break;
                }
            }

            for (int col = 3; col <= maxBalanceColumn; col++)
            {
                if (resultRow > 0)
                {
                    if (CellData(importBalanceSheet, resultRow, col) != Imported)
                    {
                        WriteCell(resultRow, col, importBalanceSheet, "");
                    }
                }
            }

            if (resultRow == 0)
            {
                resultRow = 11;
            }

            WriteCell(resultRow, 2, importBalanceSheet, "");

            // Check for duplicate part numbers in sheet
            HashSet<string> partNumbers = new HashSet<string>();
            for (int row = 11; row < resultRow; row++)
            {
                data = CellData(importBalanceSheet, row, 1);
                if (partNumbers.Contains(data.ToUpper()))
                {
                    msg = string.Format("Duplicate part number = {0}. ", data);
                    WriteCell(resultRow, 2, importBalanceSheet, msg, append: true);
                    WriteCell(validationRow++, 2, validationSheet, msg);
                }
                else
                {
                    partNumbers.Add(data.ToUpper());
                }
            }

            if (importBalanceSheet != null)
            {
                if (CellData(importBalanceSheet, 8, 1) != "" ||
                    CellData(importBalanceSheet, 9, 1) != "" ||
                    CellData(importBalanceSheet, 8, 2) != "Name:" ||
                    CellData(importBalanceSheet, 9, 2) != "Branch Or Laydown:" ||
                    CellData(importBalanceSheet, 10, 1) != "PartNumber" ||
                    CellData(importBalanceSheet, 10, 2) != "Description")
                {
                    msg = "Branch Or Laydown Balance Sheet doesn't conform to the template for this import. Check the Row Headings. ";
                    WriteCell(resultRow, 2, importBalanceSheet, msg, append: true);
                    WriteCell(validationRow++, 2, validationSheet, msg);
                }

                for (int row = 11; row < resultRow; row++)
                {
                    if (!EmptyRow(importBalanceSheet, row, maxBalanceColumn + 1)) // the last arg is the first column not checked
                    {
                        data = CellData(importBalanceSheet, row, 1);
                        if (data == "")
                        {
                            msg = "PartNumber can not be blank. ";
                            WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                        }
                        else
                        {
                            if (!ValidateSheetValue(data, catalogSheet, 1))
                            {
                                msg = "Part Number not found on Product Catalog import sheet. ";
                                WriteCell(resultRow, 2, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column 1, {1} ", row, msg), append: true);
                            }
                        }
                    }
                }

                for (int column = 3; column <= maxBalanceColumn; column++)
                {
                    if (!NoImportQuantities(importBalanceSheet, 11, resultRow, column))
                    {
                        WriteCell(resultRow, column, importBalanceSheet, "");

                        string name = CellData(importBalanceSheet, 8, column);
                        if (name == "")
                        {
                            msg = "Name is required and not found. ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 8, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 8, ColumnIndexToLetter(column), msg), append: true);
                        }

                        string branchOrLaydown = CellData(importBalanceSheet, 9, column);
                        if (branchOrLaydown == "")
                        {
                            msg = "Branch Or Laydown is required and not found. Must be \"Branch\" or \"Laydown\". ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 9, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 9, ColumnIndexToLetter(column), msg), append: true);
                        }
                        else
                        {
                            string nameKey = name + "|" + branchOrLaydown;
                            if (seenNames.Contains(nameKey))
                            {
                                msg = "Duplicate [Name, BranchOrLaydown] on this sheet is not allowed. ";
                                WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 9, msg), append: true);
                                WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 9, ColumnIndexToLetter(column), msg), append: true);
                            }
                            else
                            {
                                seenNames.Add(nameKey);
                            }
                        }

                        data = CellData(importBalanceSheet, 10, column);
                        if (data != "")
                        {
                            msg = "Row 10 should be blank. ";
                            WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", 12, msg), append: true);
                            WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", 12, ColumnIndexToLetter(column), msg), append: true);
                        }

                        for (int row = 11; row <= maxBalanceRow; row++)
                        {
                            data = CellData(importBalanceSheet, row, column);
                            if (data != "")
                            {
                                bool isNumber = double.TryParse(data, out double tempDouble);
                                if (!isNumber)
                                {
                                    msg = "A non-blank quantity must be a valid number. ";
                                    WriteCell(resultRow, column, importBalanceSheet, string.Format("Row {0}, {1}  ", row, msg), append: true);
                                    WriteCell(validationRow++, 2, validationSheet, string.Format("Row {0}, Column {1}, {2} ", row, ColumnIndexToLetter(column), msg), append: true);
                                }
                            }
                        }
                    }
                    else
                    {
                        msg = "There are no quantities to Import. ";
                        WriteCell(resultRow, column, importBalanceSheet, string.Format("{0}  ", msg), append: true);
                        WriteCell(validationRow++, 2, validationSheet, string.Format("Column {0}, {1} ", ColumnIndexToLetter(column), msg), append: true);
                    }
                }
            }
        }

        #endregion

        #region Validation Helpers

        /// <summary>
        /// Look for data equal to a value in column of lookupSheet returning found = true or false
        /// if listName or listNameColumn are provided, they must correspond to a column containing a listname
        /// corresponding to the value.  For example listName can be "Manufacturer" in column 1 of Global List Import
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="lookupSheet"></param>
        /// <param name="column"></param>
        /// <param name="listName"> Optional listName for data</param>
        /// <param name="listNameColumn">Optional column for listName</param>
        /// <returns>true if validated</returns>
        internal bool ValidateSheetValue(string data, Excel.Worksheet lookupSheet, int column,
        string value1 = "", int value1Column = 0,
        string value2 = "", int value2Column = 0,
        string value3 = "", int value3Column = 0,
        bool norm1 = false,
        bool norm2 = false,
        bool norm3 = false)
        {
            bool found = false;
            if (lookupSheet == null)
            {
                return false;      // Exit point
            }

            var activeRange = (object[,])lookupSheet.UsedRange.Value;
            int maxRows = activeRange.GetLength(0);

            for (int row = 7; row <= maxRows; row++)
            {
                bool checkData = CellData(lookupSheet, row, column).ToUpper() == data.ToUpper();
                bool check1 = value1Column == 0 || (norm1 && NormStr(CellData(lookupSheet, row, value1Column)) == NormStr(value1))
                    || (!norm1 && CellData(lookupSheet, row, value1Column).ToUpper() == value1.ToUpper());
                bool check2 = value2Column == 0 || (norm2 && NormStr(CellData(lookupSheet, row, value2Column)) == NormStr(value2))
                    || (!norm2 && CellData(lookupSheet, row, value2Column).ToUpper() == value2.ToUpper());
                bool check3 = value3Column == 0 || (norm3 && NormStr(CellData(lookupSheet, row, value3Column)) == NormStr(value3))
                    || (!norm3 && CellData(lookupSheet, row, value3Column).ToUpper() == value3.ToUpper());

                if (checkData && check1 && check2 && check3)
                {
                    found = true;
                    break;
                }
            }
            return found;
        }

        internal Excel.Worksheet FindSheet(string catalogIdentifier, Excel.Workbook wb, Dictionary<string, string> sheetDic)
        {
            if (sheetDic.ContainsKey(catalogIdentifier))
            {
                foreach (Excel.Worksheet sh in wb.Sheets)
                {
                    if (sh.Name == sheetDic[catalogIdentifier])
                    {
                        return sh;
                    }
                }
            }
            return null;
        }

        internal string CustNameFromJobSiteName(string jobSiteName, Excel.Worksheet jobSiteSheet)
        {
            string foundCustName = "";
            var activeRange = (object[,])jobSiteSheet.UsedRange.Value;
            int maxRows = activeRange.GetLength(0);

            for (int row = 7; row <= maxRows; row++)
            {

                if (CellData(jobSiteSheet, row, 3).ToUpper() == jobSiteName.ToUpper())
                {
                    foundCustName = CellData(jobSiteSheet, row, 2).ToUpper();
                    break;
                }
            }
            return foundCustName;
        }

        /// <summary>
        /// For non-balance worksheets, returns true if identified row has no non-blank values. 
        ///  
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="resultCol"></param>
        /// <returns></returns>
        internal bool EmptyRow(Excel.Worksheet ws, int row, int resultCol)
        {
            for (int col = 1; col < resultCol; col++)
            {
                if (CellData(ws, row, col).Trim() != "")
                {
                    return false;
                }
            }
            return true;
        }


        internal bool NoImportQuantities(Excel.Worksheet ws, int row1, int resultRow, int column)
        {
            for (int row = row1; row < resultRow; row++)
            {
                if (CellData(ws, row, column) != "")
                {
                    return false;
                }
            }
            return true;
        }
        #endregion

        #region Create Templates
        internal void CreateTemplates(string versionStr)
        {
            // If workbook or sheet is null then exit
            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null ||
                Globals.QuantifyImporter.Application.ActiveSheet == null)
            {
                return;
            }

            if (Globals.QuantifyImporter.Application.ActiveWorkbook == null)
            {
                MessageBoxHelper.Show("This workbook is currently in protected mode. Please enable editing to continue.", "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Excel.Workbook wb = Globals.QuantifyImporter.Application.ActiveWorkbook;
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.QuantifyImporter.Application.ActiveSheet;

            var configSeq = ImportConfig.GetAllImportConfigs().OrderBy(x => x.SheetOrder);
            foreach (ImportConfig config in configSeq)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                try
                {
                    ws.Name = config.Name;
                }
                catch
                {
                    MessageBoxHelper.Show(string.Format("Unable to create new sheet named {0}.  May be a duplicate name.", config.Name), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                ws.Range["A1", "D1"].Merge();
                ws.Range["A1", "D1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                ws.Range["A1", "D1"].RowHeight = 68;
                WriteCell(1, 1, ws, "CompanyName", requiredField: true);

                WriteCell(3, 1, ws, config.ImportIdentifier);
                WriteCell(4, 1, ws, string.Format("Template vrs {0} ", versionStr));
                int row = 6;
                int col = 1;

                // Unprotected cells
                foreach (var importCol in config.ImportColumns)
                {
                    WriteCell(row, col, ws, importCol.ColumnHeading, false, requiredField: importCol.IsRequired);
                    ((Excel.Range)ws.Cells[row, col]).Columns.AutoFit();

                    col += 1;
                }
            }


            // Create Template for Scaffold Tag Balance Importer
            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            string configName = "Scaffold Tag Balances";
            try
            {
                ws1.Name = configName;
            }
            catch
            {
                MessageBoxHelper.Show(string.Format("Unable to create new sheet named {0}.  May be a duplicate name.", configName), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ws1.Range["A1", "D1"].Merge();
            ws1.Range["A1", "D1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            ws1.Range["A1", "D1"].RowHeight = 68;
            WriteCell(1, 1, ws1, "CompanyName", requiredField: true);

            WriteCell(3, 1, ws1, "Import: " + configName);
            WriteCell(4, 1, ws1, string.Format("Template vrs {0}", versionStr));
            WriteCell(7, 2, ws1, "Rent Start:", requiredField: true);
            WriteCell(8, 2, ws1, "Job Site:", requiredField: true);
            WriteCell(9, 2, ws1, "Tag:", requiredField: true);
            WriteCell(10, 2, ws1, "Activity:", requiredField: false);
            WriteCell(11, 2, ws1, "Summary:", requiredField: false);
            WriteCell(12, 2, ws1, "Is Billable:", requiredField: true);
            WriteCell(13, 1, ws1, "PartNumber", requiredField: true);
            WriteCell(13, 2, ws1, "Description", requiredField: true);

            // Set border for 1st tag
            for (int row = 14; row <= 19; row++)
            {
                ((Excel.Range)ws1.Cells[row, 3]).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
            }

            // Create Template for Job Balance Importer
            ws1 = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            configName = "Job Site Balances";
            try
            {
                ws1.Name = configName;
            }
            catch
            {
                MessageBoxHelper.Show(string.Format("Unable to create new sheet named {0}.  May be a duplicate name.", configName), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ws1.Range["A1", "D1"].Merge();
            ws1.Range["A1", "D1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            ws1.Range["A1", "D1"].RowHeight = 68;
            WriteCell(1, 1, ws1, "CompanyName", requiredField: true);

            WriteCell(3, 1, ws1, "Import: " + configName);
            WriteCell(4, 1, ws1, string.Format("Template vrs {0}", versionStr));
            WriteCell(7, 2, ws1, "Rent Start:", requiredField: true);
            WriteCell(8, 2, ws1, "Job Site:", requiredField: true);

            WriteCell(9, 2, ws1, "Is Billable:", requiredField: true);
            WriteCell(10, 1, ws1, "PartNumber", requiredField: true);
            WriteCell(10, 2, ws1, "Description", requiredField: true);

            // Set border for 1st tag
            for (int row = 10; row <= 13; row++)
            {
                ((Excel.Range)ws1.Cells[row, 3]).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
            }

            // Create Template for Branch Or Laydown Balance Importer
            ws1 = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            configName = "Branch Or Laydown Balances";
            try
            {
                ws1.Name = configName;
            }
            catch
            {
                MessageBoxHelper.Show(string.Format("Unable to create new sheet named {0}.  May be a duplicate name.", configName), "Quantify Excel Importer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ws1.Range["A1", "D1"].Merge();
            ws1.Range["A1", "D1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            ws1.Range["A1", "D1"].RowHeight = 68;
            WriteCell(1, 1, ws1, "CompanyName", requiredField: true);

            WriteCell(3, 1, ws1, "Import: " + configName);
            WriteCell(4, 1, ws1, string.Format("Template vrs {0}", versionStr));

            WriteCell(8, 2, ws1, "Name:", requiredField: true);
            WriteCell(9, 2, ws1, "Branch Or Laydown:", requiredField: true);
            WriteCell(10, 1, ws1, "PartNumber", requiredField: true);
            WriteCell(10, 2, ws1, "Description", requiredField: true);


            // Set border for 1st tag
            for (int row = 11; row <= 18; row++)
            {
                ((Excel.Range)ws1.Cells[row, 3]).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
            }

        }

        #endregion

        #region Helpers

        // See EmptyRow in another region
        internal bool EmptyCol(Excel.Worksheet ws, int col)
        {
            var activeRange = (object[,])ws.UsedRange.Value;

            for (int row = 1; row < activeRange.GetLength(0); row++)
            {
                if (CellData(ws, row, col).Trim() != "")
                {
                    return false;
                }
            }
            return true;
        }

        private Address GetPartnerAddress(BusinessPartner partner, AddressTypes type)
        {
            Address returnAddress = null;

            bool found = false;

            foreach (Address address in partner.Addresses)
            {
                if (address.AddressType == type)
                {
                    found = true;
                    returnAddress = address;
                    break;
                }
            }

            if (!found)
            {
                returnAddress = Address.NewAddress();
                returnAddress.AddressType = AddressTypes.Business;
                partner.Addresses.Add(returnAddress);
            }

            return returnAddress;
        }

        private Address GetLocationAddress(StockingLocation location, AddressTypes type)
        {
            Address returnAddress = null;

            bool found = false;

            foreach (Address address in location.Addresses)
            {
                if (address.AddressType == type)
                {
                    found = true;
                    returnAddress = address;
                    break;
                }
            }

            if (!found)
            {
                returnAddress = Address.NewAddress();
                returnAddress.AddressType = AddressTypes.Business;
                location.Addresses.Add(returnAddress);
            }

            return returnAddress;
        }

        private BoolResult CellBool(Tuple<string, ImportConfig.ImportColunnType, int> data, bool allowEmpty = false, bool allowTrueFalse = true)
        {
            return CellBool(data.Item1, allowEmpty, allowTrueFalse);
        }

        private BoolResult CellBool(string data, bool allowEmpty = false, bool allowTrueFalse = true)
        {
            BoolResult result = BoolResult.Invalid;

            if (data != null && data != "")
            {
                switch (data.ToUpper())
                {
                    case "TRUE":
                        result = BoolResult.True;
                        break;
                    case "FALSE":
                        result = BoolResult.False;
                        break;
                    case "YES":
                        result = allowTrueFalse ? BoolResult.True : BoolResult.Invalid;
                        break;
                    case "NO":
                        result = allowTrueFalse ? BoolResult.False : BoolResult.Invalid;
                        break;
                    default:
                        break;

                }
            }
            else
            {
                result = allowEmpty ? BoolResult.Empty : BoolResult.Invalid;
            }
            return result;
        }

        private Excel.Worksheet GetValidationSheet(List<string> sheetNames, Excel.Workbook wb, string prefix = "Validation")
        {
            string prefixU = prefix.ToUpper();
            int prefixLength = prefixU.Length;
            int maxSuffix = -1;
            int suffix = 0;
            foreach (string nameStr in sheetNames)
            {
                if (nameStr.Left(prefixLength).ToUpper() == prefixU)
                {
                    if (nameStr.Length > prefixLength)
                    {
                        int.TryParse(nameStr.Substring(prefixLength), out suffix);
                        maxSuffix = maxSuffix < suffix ? suffix : maxSuffix;
                    }
                    else
                    {
                        maxSuffix = maxSuffix == -1 ? 0 : maxSuffix;
                    }
                }
            }
            string validationSheetName = prefix + (maxSuffix >= 0 ? (maxSuffix + 1).ToString() : "");
            wb.Sheets.Add(Before: wb.Sheets[1]);
            Excel.Worksheet sheet = (Excel.Worksheet)wb.Sheets[1];
            sheet.Name = validationSheetName;
            return sheet;
        }

        private string GetSheetIdentifier(string name, Excel.Workbook wb, IOrderedEnumerable<ImportConfig> configs)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.ToUpper() == name.ToUpper())
                {
                    string signature = CellData(sheet, 3, 1);
                    string normSig = NormStr(signature);

                    switch (normSig)
                    {
                        case "IMPORT:BRANCHORLAYDOWNBALANCES":
                            return "Import: Branch Or Laydown Balances";

                        case "IMPORT:SCAFFOLDTAGBALANCES":
                            return "Import: Scaffold Tag Balances";

                        case "IMPORT:JOBSITEBALANCES":
                            return "Import: Job Site Balances";
                        default:
                            foreach (var config in configs)
                            {
                                if (NormStr(signature) == NormStr(config.ImportIdentifier))
                                {
                                    return config.ImportIdentifier;
                                }
                            }
                            return NonImportSheetID;
                    }
                }
            }
            return NonImportSheetID;
        }

        private bool CheckWorksheetHeaders(Excel.Worksheet ws, int maxRows, int maxColumns, string importIdentifier, string resultHeader,
            int resultCol, out string headerIssues, out ImportConfig importConfig)
        {
            importConfig = ImportConfig.GetImportConfig(NormStr(importIdentifier));
            headerIssues = VerifyRequiredHeaders(ws, importConfig, maxRows, maxColumns, importIdentifier, resultHeader);
            RemoveColumnBackgroundColors(ws, importConfig, HeaderRow, maxRows, resultCol);
            return headerIssues != "";
        }

        private Excel.Worksheet GetSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.ToUpper() == name.ToUpper())
                {
                    return sheet;
                }
            }

            return null;

        }

        /// <summary>
        /// Calcualtes a specific day of a month based on parameters
        /// Returns false if it fails
        /// </summary>
        /// <param name="date"> The month containing date is the month calculated</param>
        /// <param name="monthlyType">First Day, Last Day, Last Sunday or Specific Day</param>
        /// <param name="dayOfMonth">An integer day of month for the Specific Day type</param>
        /// <param name="calculatedDate">The output date if successful, else the date</param>
        /// <returns>true if successful, else false</returns>
        private bool CalculatedDateInMonth(DateTime date, MonthlyBillingCycleType monthlyType, int dayOfMonth, out DateTime calculatedDate)
        {
            switch (monthlyType)
            {
                case MonthlyBillingCycleType.LastDay:
                    calculatedDate = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
                    return true;
                case MonthlyBillingCycleType.FirstDay:
                    calculatedDate = new DateTime(date.Year, date.Month, 1);
                    return true;
                case MonthlyBillingCycleType.LastSunday:
                    calculatedDate = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
                    while (calculatedDate.DayOfWeek != DayOfWeek.Sunday)
                    {
                        calculatedDate = calculatedDate.AddDays(-1);
                    }
                    return true;
                case MonthlyBillingCycleType.SpecificDay:
                    int lastDay = DateTime.DaysInMonth(date.Year, date.Month);
                    lastDay = dayOfMonth < lastDay ? dayOfMonth : lastDay;
                    if (lastDay > 0)
                    {
                        calculatedDate = new DateTime(date.Year, date.Month, lastDay);
                        return true;
                    }
                    else
                    {
                        calculatedDate = date;
                        return false;
                    }
                default:
                    calculatedDate = date;
                    return false;
            }
        }
        #endregion

        #region Cell Contents and Color Changes

        /// <summary>
        /// Return the column letter for column with column n 
        /// </summary>
        /// <param name="col"></param>
        ///     column index where column A is 1
        ///     col is valie from 1 to 702 == "ZZ"
        /// <returns></returns>
        string ColumnIndexToLetter(int col)
        {
            const int ColumnZZ = 702;
            string alpha = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (col > ColumnZZ)
            {
                return "col > ZZ";
            }
            if (col < 1)
            {
                return "col < A";
            }
            int x = ((col - 1) / 26);
            int y = (col % 26 == 0 ? 26 : col % 26);

            return (alpha[x].ToString() + alpha[y].ToString()).Trim();
        }

        /// <summary>
        /// Get details for a particular column in the worksheet (specified by the normalized header value)
        /// </summary>
        /// <param name="header"></param>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="importDict"></param>
        /// <returns></returns>

        private Tuple<string, ImportConfig.ImportColunnType, int>
            CellData(string header, Excel.Worksheet ws, int row, Dictionary<string, ImportConfig.ImportColumn> importDict)
        {
            ImportConfig.ImportColumn importCol = null;
            bool found = importDict.TryGetValue(header, out importCol);
            if (!found || importCol == null)
            {
                return null;
            }
            var range = ((Excel.Range)ws.Cells[row, importCol.Column]);
            string value = ((string)range.Formula).Trim() == "" ? "" : range.Value.ToString().Trim();
            ImportConfig.ImportColunnType type = importCol.ColumnType;
            int column = importCol.Column;
            return new Tuple<string, ImportConfig.ImportColunnType, int>(value, type, column);
        }


        private string CellData(Excel.Worksheet ws, int row, int col)
        {

            var range = ((Excel.Range)ws.Cells[row, col]);
            string value = ((string)range.Formula).Trim() == "" ? "" : range.Value.ToString().Trim();
            return value;
        }

        /// <summary>
        /// Writes and/or changes Interior Color for a cell; can change color of a 2nd column and other stuff
        /// </summary>
        /// <param name="row">Target row of cell to change</param>
        /// <param name="col">Target col of cell to change</param>
        /// <param name="ws">The worksheet to modify</param>
        /// <param name="value"> Optional text to write; defaults to no change</param>
        /// <param name="useColor">If true, then Interior color is changed for columns specified.  If color is null, it is removed.</param>
        /// <param name="append">Option to append rather than replace text value</param>
        /// <param name="underline">Option to underline the cell)</param>
        /// <param name="bold"> Option to bold the cell</param>
        /// <param name="color">Optional Color for col and dataCol, if it is non-null</param>
        /// <param name="dataRow">Optional 2nd row to change Interior Color</param> 
        /// <param name="dataCol">Optional 2nd column to change Interior Color</param>
        /// <param name="requiredField">Optional uses Required colors and bolds for the heading</param>
        private void WriteCell(int row, int col, Excel.Worksheet ws, string value = null, bool useColor = true, bool append = false,
            bool underline = false, bool bold = false, System.Drawing.Color? color = null, int? dataRow = null, int? dataCol = null,
            bool requiredField = false, string numberformat = "")
        {
            if (value != null)
            {
                if (append)
                {
                    ((Excel.Range)ws.Cells[row, col]).Value += value;
                }
                else
                {
                    ((Excel.Range)ws.Cells[row, col]).Value = value;
                }
            }

            if (useColor)
            {
                if (color != null)
                {
                    ((Excel.Range)ws.Cells[row, col]).Interior.Color = color;
                    if (dataRow != null && dataCol != null)
                    {
                        ((Excel.Range)ws.Cells[dataRow, dataCol]).Interior.Color = color;
                    }
                    else if (dataCol != null)
                    {
                        ((Excel.Range)ws.Cells[row, dataCol]).Interior.Color = color;
                    }
                    else if (dataRow != null)
                    {
                        ((Excel.Range)ws.Cells[dataRow, col]).Interior.Color = color;
                    }
                }
                else
                {
                    ((Excel.Range)ws.Cells[row, col]).Interior.ColorIndex = 0;

                    if (dataRow != null && dataCol != null)
                    {
                        ((Excel.Range)ws.Cells[dataRow, dataCol]).Interior.ColorIndex = 0;
                    }
                    else if (dataCol != null)
                    {
                        ((Excel.Range)ws.Cells[row, dataCol]).Interior.ColorIndex = 0;
                    }
                    else if (dataRow != null)
                    {
                        ((Excel.Range)ws.Cells[dataRow, col]).Interior.ColorIndex = 0;
                    }
                }
            }

            if (underline)
            {
                ((Excel.Range)ws.Cells[row, col]).Font.Underline = true;
            }
            if (bold)
            {
                ((Excel.Range)ws.Cells[row, col]).Font.Bold = true;
            }

            if (requiredField)
            {
                ((Excel.Range)ws.Cells[row, col]).Interior.Color = this.requiredInterior;
                ((Excel.Range)ws.Cells[row, col]).Font.Color = this.requiredFontColor;
                ((Excel.Range)ws.Cells[row, col]).Font.Bold = true;
            }

            if (numberformat != "")
            {
                ((Excel.Range)ws.Cells[row, col]).NumberFormat = numberformat;
            }
        }

        /// <summary>
        /// Removes background fill from columns that are to be imported
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="config"></param>
        /// <param name="headerRow"></param>
        /// <param name="maxRow"></param>
        /// <returns></returns>
        private bool RemoveColumnBackgroundColors(Excel.Worksheet ws, ImportConfig config, int headerRow, int maxRow, int resultCol)
        {
            for (var row = headerRow + 1; row <= maxRow; row++)
            {
                var prevResult = ((Excel.Range)ws.Cells[row, resultCol]).Value;
                if (prevResult != null)
                {
                    if (prevResult.ToString() != "Imported.")
                    {
                        ((Excel.Range)ws.Cells[row, resultCol]).ClearContents();
                    }
                    foreach (var importCol in config.ImportColumns)
                    {
                        if (importCol.Column >= 0)
                            ((Excel.Range)ws.Cells[row, importCol.Column + 1]).Interior.ColorIndex = 0;
                    }
                }
            }
            return true;
        }

        private string savedE1 = "";
        private bool progressStarted = false;
        private const int progressRow = 2;
        private const int progressCol = 1;

        private void ClearProgress(Excel.Worksheet ws)
        {
            // WriteCell(progressRow, progressCol, ws, savedE1);
            savedE1 = "";
            progressStarted = false;
        }

        private void ShowProgress(Excel.Worksheet ws, ProgressDisplayType displayType, int currentValue, int maxValue, string version, int resultColumn = -1)
        {
            if (!progressStarted)
            {
                savedE1 = (string)((Excel.Range)ws.Cells[progressRow, progressCol]).Text;
                progressStarted = true;
            }
            WriteCell(progressRow, progressCol, ws, string.Format("Progress - {0}: {1} of {2} {3}",
                displayType.ToString(), currentValue, maxValue, version));

            // Display the count over the result column for row based imports (not balances)
            if (resultColumn != -1)
            {
                WriteCell(progressRow, resultColumn, ws, string.Format("{0} of {1}",
                     currentValue, maxValue));
            }
        }
        #endregion
    }
}