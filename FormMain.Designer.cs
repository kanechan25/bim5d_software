
namespace QSKSKS
{
    partial class FormMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.bImport = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bSupport = new System.Windows.Forms.Button();
            this.bBOQ = new System.Windows.Forms.Button();
            this.bClearProj = new System.Windows.Forms.Button();
            this.bRebar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // bImport
            // 
            this.bImport.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bImport.Location = new System.Drawing.Point(12, 13);
            this.bImport.Name = "bImport";
            this.bImport.Size = new System.Drawing.Size(229, 26);
            this.bImport.TabIndex = 0;
            this.bImport.Text = "Import Files";
            this.bImport.UseVisualStyleBackColor = true;
            this.bImport.Click += new System.EventHandler(this.bImport_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(276, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(231, 66);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(276, 102);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(146, 22);
            this.label1.TabIndex = 2;
            this.label1.Text = "LET\'S JOIN US";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(344, 136);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(163, 22);
            this.label2.TabIndex = 3;
            this.label2.Text = " && RIDE ON B.I.M";
            // 
            // bSupport
            // 
            this.bSupport.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bSupport.Location = new System.Drawing.Point(12, 57);
            this.bSupport.Name = "bSupport";
            this.bSupport.Size = new System.Drawing.Size(229, 26);
            this.bSupport.TabIndex = 4;
            this.bSupport.Text = "Support Cubicost";
            this.bSupport.UseVisualStyleBackColor = true;
            this.bSupport.Click += new System.EventHandler(this.bSupport_Click);
            // 
            // bBOQ
            // 
            this.bBOQ.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bBOQ.Location = new System.Drawing.Point(12, 146);
            this.bBOQ.Name = "bBOQ";
            this.bBOQ.Size = new System.Drawing.Size(229, 26);
            this.bBOQ.TabIndex = 6;
            this.bBOQ.Text = "BOQ Form";
            this.bBOQ.UseVisualStyleBackColor = true;
            this.bBOQ.Click += new System.EventHandler(this.bBOQ_Click);
            // 
            // bClearProj
            // 
            this.bClearProj.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bClearProj.Location = new System.Drawing.Point(12, 188);
            this.bClearProj.Name = "bClearProj";
            this.bClearProj.Size = new System.Drawing.Size(229, 26);
            this.bClearProj.TabIndex = 5;
            this.bClearProj.Text = "Clear Project";
            this.bClearProj.UseVisualStyleBackColor = true;
            this.bClearProj.Click += new System.EventHandler(this.bClearProj_Click);
            // 
            // bRebar
            // 
            this.bRebar.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bRebar.Location = new System.Drawing.Point(12, 102);
            this.bRebar.Name = "bRebar";
            this.bRebar.Size = new System.Drawing.Size(229, 26);
            this.bRebar.TabIndex = 7;
            this.bRebar.Text = "Support Rebar";
            this.bRebar.UseVisualStyleBackColor = true;
            this.bRebar.Click += new System.EventHandler(this.bRebar_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(530, 226);
            this.Controls.Add(this.bRebar);
            this.Controls.Add(this.bBOQ);
            this.Controls.Add(this.bClearProj);
            this.Controls.Add(this.bSupport);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.bImport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = " TTD BIM 5D";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormMain_Closed);
            this.Load += new System.EventHandler(this.FormMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bImport;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button bSupport;
        private System.Windows.Forms.Button bBOQ;
        private System.Windows.Forms.Button bClearProj;
        private System.Windows.Forms.Button bRebar;
    }
}