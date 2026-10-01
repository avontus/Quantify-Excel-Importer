namespace Avontus.Quantify.ExcelImporter
{
    partial class frmInformationList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmInformationList));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pnlMain = new System.Windows.Forms.Panel();
            this.btnCopy = new System.Windows.Forms.Button();
            this.statStrip = new System.Windows.Forms.StatusStrip();
            this.lblCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.dgErrors = new System.Windows.Forms.DataGridView();
            this.colImage = new System.Windows.Forms.DataGridViewImageColumn();
            this.colMessage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bsInfoList = new System.Windows.Forms.BindingSource(this.components);
            this.pnlMain.SuspendLayout();
            this.statStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgErrors)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsInfoList)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlMain
            // 
            this.pnlMain.Controls.Add(this.btnCopy);
            this.pnlMain.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlMain.Location = new System.Drawing.Point(0, 0);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new System.Drawing.Size(591, 31);
            this.pnlMain.TabIndex = 7;
            // 
            // btnCopy
            // 
            this.btnCopy.Image = ((System.Drawing.Image)(resources.GetObject("btnCopy.Image")));
            this.btnCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCopy.Location = new System.Drawing.Point(3, 2);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(69, 27);
            this.btnCopy.TabIndex = 3;
            this.btnCopy.Text = "Copy";
            this.btnCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // statStrip
            // 
            this.statStrip.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblCount});
            this.statStrip.Location = new System.Drawing.Point(0, 280);
            this.statStrip.Name = "statStrip";
            this.statStrip.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.statStrip.Size = new System.Drawing.Size(591, 22);
            this.statStrip.TabIndex = 6;
            this.statStrip.Text = "StatusStrip1";
            // 
            // lblCount
            // 
            this.lblCount.Name = "lblCount";
            this.lblCount.Size = new System.Drawing.Size(83, 17);
            this.lblCount.Text = "[count of items]";
            // 
            // dgErrors
            // 
            this.dgErrors.AllowUserToAddRows = false;
            this.dgErrors.AllowUserToDeleteRows = false;
            this.dgErrors.AllowUserToResizeColumns = false;
            this.dgErrors.AllowUserToResizeRows = false;
            this.dgErrors.AutoGenerateColumns = false;
            this.dgErrors.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgErrors.ColumnHeadersVisible = false;
            this.dgErrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colImage,
            this.colMessage});
            this.dgErrors.DataSource = this.bsInfoList;
            this.dgErrors.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgErrors.Location = new System.Drawing.Point(0, 31);
            this.dgErrors.Name = "dgErrors";
            this.dgErrors.ReadOnly = true;
            this.dgErrors.RowHeadersVisible = false;
            this.dgErrors.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgErrors.Size = new System.Drawing.Size(591, 249);
            this.dgErrors.TabIndex = 8;
            // 
            // colImage
            // 
            this.colImage.DataPropertyName = "TheImage";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.NullValue = ((object)(resources.GetObject("dataGridViewCellStyle1.NullValue")));
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black;
            this.colImage.DefaultCellStyle = dataGridViewCellStyle1;
            this.colImage.HeaderText = "Image";
            this.colImage.MinimumWidth = 18;
            this.colImage.Name = "colImage";
            this.colImage.ReadOnly = true;
            this.colImage.Width = 18;
            // 
            // colMessage
            // 
            this.colMessage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colMessage.DataPropertyName = "TheMessage";
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            this.colMessage.DefaultCellStyle = dataGridViewCellStyle2;
            this.colMessage.HeaderText = "TheMessage";
            this.colMessage.Name = "colMessage";
            this.colMessage.ReadOnly = true;
            // 
            // bsInfoList
            // 
            this.bsInfoList.DataSource = typeof(Avontus.Quantify.ExcelImporter.InformationList);
            // 
            // frmInformationList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 302);
            this.Controls.Add(this.dgErrors);
            this.Controls.Add(this.pnlMain);
            this.Controls.Add(this.statStrip);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmInformationList";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Import Details";
            this.Load += new System.EventHandler(this.frmInformationList_Load);
            this.pnlMain.ResumeLayout(false);
            this.statStrip.ResumeLayout(false);
            this.statStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgErrors)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bsInfoList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.BindingSource bsInfoList;
        internal System.Windows.Forms.Panel pnlMain;
        internal System.Windows.Forms.Button btnCopy;
        internal System.Windows.Forms.StatusStrip statStrip;
        internal System.Windows.Forms.ToolStripStatusLabel lblCount;
        private System.Windows.Forms.DataGridView dgErrors;
        private System.Windows.Forms.DataGridViewImageColumn colImage;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMessage;
    }
}