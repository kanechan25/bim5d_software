
namespace QSKSKS
{
    partial class FormLoadExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormLoadExcel));
            this.dgvExcel = new System.Windows.Forms.DataGridView();
            this.cmtInsert = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.insertQuantityHereToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tbSelectRow = new System.Windows.Forms.TextBox();
            this.bOK = new System.Windows.Forms.Button();
            this.tbCategory = new System.Windows.Forms.TextBox();
            this.bGet = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).BeginInit();
            this.cmtInsert.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvExcel
            // 
            this.dgvExcel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvExcel.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvExcel.ContextMenuStrip = this.cmtInsert;
            this.dgvExcel.Location = new System.Drawing.Point(0, 47);
            this.dgvExcel.MultiSelect = false;
            this.dgvExcel.Name = "dgvExcel";
            this.dgvExcel.Size = new System.Drawing.Size(672, 364);
            this.dgvExcel.TabIndex = 0;
            this.dgvExcel.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvExcel_CellContentClick);
            this.dgvExcel.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvExcel_CellContentDoubleClick);
            // 
            // cmtInsert
            // 
            this.cmtInsert.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.insertQuantityHereToolStripMenuItem});
            this.cmtInsert.Name = "cmtInsert";
            this.cmtInsert.Size = new System.Drawing.Size(181, 26);
            this.cmtInsert.Opening += new System.ComponentModel.CancelEventHandler(this.cmtInsert_Opening);
            // 
            // insertQuantityHereToolStripMenuItem
            // 
            this.insertQuantityHereToolStripMenuItem.Name = "insertQuantityHereToolStripMenuItem";
            this.insertQuantityHereToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.insertQuantityHereToolStripMenuItem.Text = "Insert Quantity Here";
            // 
            // tbSelectRow
            // 
            this.tbSelectRow.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbSelectRow.Location = new System.Drawing.Point(7, 12);
            this.tbSelectRow.Name = "tbSelectRow";
            this.tbSelectRow.Size = new System.Drawing.Size(48, 22);
            this.tbSelectRow.TabIndex = 1;
            // 
            // bOK
            // 
            this.bOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.bOK.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bOK.Location = new System.Drawing.Point(581, 11);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(84, 24);
            this.bOK.TabIndex = 302;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = true;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // tbCategory
            // 
            this.tbCategory.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbCategory.Location = new System.Drawing.Point(63, 12);
            this.tbCategory.Name = "tbCategory";
            this.tbCategory.Size = new System.Drawing.Size(147, 22);
            this.tbCategory.TabIndex = 303;
            // 
            // bGet
            // 
            this.bGet.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGet.Location = new System.Drawing.Point(216, 11);
            this.bGet.Name = "bGet";
            this.bGet.Size = new System.Drawing.Size(155, 24);
            this.bGet.TabIndex = 305;
            this.bGet.Text = "Get All Heading Rows";
            this.bGet.UseVisualStyleBackColor = true;
            this.bGet.Click += new System.EventHandler(this.bGet_Click);
            // 
            // FormLoadExcel
            // 
            this.AcceptButton = this.bOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(672, 411);
            this.Controls.Add(this.bGet);
            this.Controls.Add(this.tbCategory);
            this.Controls.Add(this.bOK);
            this.Controls.Add(this.tbSelectRow);
            this.Controls.Add(this.dgvExcel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormLoadExcel";
            this.Text = "Load Excel BOQ File";
            this.Load += new System.EventHandler(this.FormLoadExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcel)).EndInit();
            this.cmtInsert.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

    }

        #endregion

        private System.Windows.Forms.DataGridView dgvExcel;
        private System.Windows.Forms.TextBox tbSelectRow;
        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.TextBox tbCategory;
        private System.Windows.Forms.Button bGet;
        private System.Windows.Forms.ContextMenuStrip cmtInsert;
        private System.Windows.Forms.ToolStripMenuItem insertQuantityHereToolStripMenuItem;
    }
}