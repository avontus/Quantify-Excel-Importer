using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Avontus.Core;

namespace Avontus.Quantify.ExcelImporter
{
    public partial class frmInformationList : Form
    {
        InformationList _informationList;
        public frmInformationList(InformationList informationList)
        {
            InitializeComponent();
            _informationList = informationList;
        }

        private void frmInformationList_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            SortedBindingList<InformationListItem> sortedList = new SortedBindingList<InformationListItem>(_informationList);
            sortedList.ApplySort("SortOrder", System.ComponentModel.ListSortDirection.Ascending);
            dgErrors.DataSource = sortedList;
            lblCount.Text = "Items: " + _informationList.Count.ToString();

            dgErrors.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Cursor = Cursors.Default;
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            try
            {
               
                string contents = "";
                foreach (DataGridViewRow row in dgErrors.Rows)
                {
                    if (row.Cells[1].Value.ToString().Length > 0)
                        contents += row.Cells[1].Value.ToString() + "\r\n";
                }
                Clipboard.SetDataObject(contents);
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                MessageBoxHelper.Show("The Clipboard could not be accessed. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
