using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Avontus.Core; 

namespace Avontus.Quantify.ExcelImporter
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            btnOK.DialogResult = DialogResult.OK;
            btnCancel.DialogResult = DialogResult.Cancel;

            if (System.Diagnostics.Debugger.IsAttached)
                txtUsername.Text = "Admin";

            if (txtUsername.Text.Length > 0)
                txtPassword.Select();
        }

		public string Username
		{
			get { return txtUsername.Text; }
			set { txtUsername.Text = value; }
		}

	    public string Password
	    {
		    get { return txtPassword.Text; }
			set { txtPassword.Text = value; }
	    }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
