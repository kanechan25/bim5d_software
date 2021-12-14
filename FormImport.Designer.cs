
namespace QSKSKS
{
    partial class FormTTD
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormTTD));
            this.panelTIO = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.bGetSExpPath = new System.Windows.Forms.Button();
            this.bGetSCatePath = new System.Windows.Forms.Button();
            this.tbSExpPath = new System.Windows.Forms.TextBox();
            this.tbSCatePath = new System.Windows.Forms.TextBox();
            this.bGetAExpPath = new System.Windows.Forms.Button();
            this.bGetACatePath = new System.Windows.Forms.Button();
            this.tbAExpPath = new System.Windows.Forms.TextBox();
            this.tbACatePath = new System.Windows.Forms.TextBox();
            this.openPath = new System.Windows.Forms.OpenFileDialog();
            this.panelRebar = new System.Windows.Forms.Panel();
            this.bGetRebarPath = new System.Windows.Forms.Button();
            this.tbRebarPath = new System.Windows.Forms.TextBox();
            this.panelBOQ = new System.Windows.Forms.Panel();
            this.bGetBOQPath = new System.Windows.Forms.Button();
            this.tbBOQPath = new System.Windows.Forms.TextBox();
            this.tbFF4Path = new System.Windows.Forms.TextBox();
            this.tbFF3Path = new System.Windows.Forms.TextBox();
            this.bGetFFPath = new System.Windows.Forms.Button();
            this.tbFF2Path = new System.Windows.Forms.TextBox();
            this.tbFF1Path = new System.Windows.Forms.TextBox();
            this.tbPlu4Path = new System.Windows.Forms.TextBox();
            this.tbPlu3Path = new System.Windows.Forms.TextBox();
            this.bGetPluPath = new System.Windows.Forms.Button();
            this.tbPlu2Path = new System.Windows.Forms.TextBox();
            this.tbPlu1Path = new System.Windows.Forms.TextBox();
            this.tbEle4Path = new System.Windows.Forms.TextBox();
            this.tbEle3Path = new System.Windows.Forms.TextBox();
            this.bGetElePath = new System.Windows.Forms.Button();
            this.tbEle2Path = new System.Windows.Forms.TextBox();
            this.tbEle1Path = new System.Windows.Forms.TextBox();
            this.tbMec4Path = new System.Windows.Forms.TextBox();
            this.tbMec3Path = new System.Windows.Forms.TextBox();
            this.bGetMecPath = new System.Windows.Forms.Button();
            this.tbMec2Path = new System.Windows.Forms.TextBox();
            this.tbMec1Path = new System.Windows.Forms.TextBox();
            this.bClose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.openPath2 = new System.Windows.Forms.OpenFileDialog();
            this.bClear = new System.Windows.Forms.Button();
            this.panelTIO.SuspendLayout();
            this.panelRebar.SuspendLayout();
            this.panelBOQ.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTIO
            // 
            this.panelTIO.Controls.Add(this.panel1);
            this.panelTIO.Controls.Add(this.bGetSExpPath);
            this.panelTIO.Controls.Add(this.bGetSCatePath);
            this.panelTIO.Controls.Add(this.tbSExpPath);
            this.panelTIO.Controls.Add(this.tbSCatePath);
            this.panelTIO.Controls.Add(this.bGetAExpPath);
            this.panelTIO.Controls.Add(this.bGetACatePath);
            this.panelTIO.Controls.Add(this.tbAExpPath);
            this.panelTIO.Controls.Add(this.tbACatePath);
            this.panelTIO.Location = new System.Drawing.Point(2, 25);
            this.panelTIO.Name = "panelTIO";
            this.panelTIO.Size = new System.Drawing.Size(1200, 112);
            this.panelTIO.TabIndex = 6;
            this.panelTIO.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(1, 145);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1258, 124);
            this.panel1.TabIndex = 17;
            // 
            // bGetSExpPath
            // 
            this.bGetSExpPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetSExpPath.Location = new System.Drawing.Point(6, 80);
            this.bGetSExpPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetSExpPath.Name = "bGetSExpPath";
            this.bGetSExpPath.Size = new System.Drawing.Size(180, 23);
            this.bGetSExpPath.TabIndex = 16;
            this.bGetSExpPath.Text = "Structural Expression File";
            this.bGetSExpPath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetSExpPath.UseVisualStyleBackColor = true;
            this.bGetSExpPath.Click += new System.EventHandler(this.bGetSExpPath_Click);
            // 
            // bGetSCatePath
            // 
            this.bGetSCatePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetSCatePath.Location = new System.Drawing.Point(6, 59);
            this.bGetSCatePath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetSCatePath.Name = "bGetSCatePath";
            this.bGetSCatePath.Size = new System.Drawing.Size(180, 23);
            this.bGetSCatePath.TabIndex = 15;
            this.bGetSCatePath.Text = "Structural Category File";
            this.bGetSCatePath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetSCatePath.UseVisualStyleBackColor = true;
            this.bGetSCatePath.Click += new System.EventHandler(this.bGetSCatePath_Click);
            // 
            // tbSExpPath
            // 
            this.tbSExpPath.Enabled = false;
            this.tbSExpPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbSExpPath.Location = new System.Drawing.Point(188, 81);
            this.tbSExpPath.Multiline = true;
            this.tbSExpPath.Name = "tbSExpPath";
            this.tbSExpPath.Size = new System.Drawing.Size(900, 21);
            this.tbSExpPath.TabIndex = 14;
            // 
            // tbSCatePath
            // 
            this.tbSCatePath.Enabled = false;
            this.tbSCatePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbSCatePath.Location = new System.Drawing.Point(188, 60);
            this.tbSCatePath.Multiline = true;
            this.tbSCatePath.Name = "tbSCatePath";
            this.tbSCatePath.Size = new System.Drawing.Size(900, 22);
            this.tbSCatePath.TabIndex = 13;
            // 
            // bGetAExpPath
            // 
            this.bGetAExpPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetAExpPath.Location = new System.Drawing.Point(6, 29);
            this.bGetAExpPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetAExpPath.Name = "bGetAExpPath";
            this.bGetAExpPath.Size = new System.Drawing.Size(180, 23);
            this.bGetAExpPath.TabIndex = 9;
            this.bGetAExpPath.Text = "Architectural Expression File";
            this.bGetAExpPath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetAExpPath.UseVisualStyleBackColor = true;
            this.bGetAExpPath.Click += new System.EventHandler(this.bGetAExpPath_Click);
            // 
            // bGetACatePath
            // 
            this.bGetACatePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetACatePath.Location = new System.Drawing.Point(6, 7);
            this.bGetACatePath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetACatePath.Name = "bGetACatePath";
            this.bGetACatePath.Size = new System.Drawing.Size(180, 23);
            this.bGetACatePath.TabIndex = 8;
            this.bGetACatePath.Text = "Architectural Category File";
            this.bGetACatePath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetACatePath.UseVisualStyleBackColor = true;
            this.bGetACatePath.Click += new System.EventHandler(this.bGetACatePath_Click);
            // 
            // tbAExpPath
            // 
            this.tbAExpPath.Enabled = false;
            this.tbAExpPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbAExpPath.Location = new System.Drawing.Point(188, 29);
            this.tbAExpPath.Multiline = true;
            this.tbAExpPath.Name = "tbAExpPath";
            this.tbAExpPath.Size = new System.Drawing.Size(900, 21);
            this.tbAExpPath.TabIndex = 7;
            // 
            // tbACatePath
            // 
            this.tbACatePath.Enabled = false;
            this.tbACatePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbACatePath.Location = new System.Drawing.Point(188, 8);
            this.tbACatePath.Multiline = true;
            this.tbACatePath.Name = "tbACatePath";
            this.tbACatePath.Size = new System.Drawing.Size(900, 22);
            this.tbACatePath.TabIndex = 6;
            // 
            // panelRebar
            // 
            this.panelRebar.Controls.Add(this.bGetRebarPath);
            this.panelRebar.Controls.Add(this.tbRebarPath);
            this.panelRebar.Location = new System.Drawing.Point(2, 134);
            this.panelRebar.Name = "panelRebar";
            this.panelRebar.Size = new System.Drawing.Size(1200, 33);
            this.panelRebar.TabIndex = 7;
            // 
            // bGetRebarPath
            // 
            this.bGetRebarPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetRebarPath.Location = new System.Drawing.Point(6, 6);
            this.bGetRebarPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetRebarPath.Name = "bGetRebarPath";
            this.bGetRebarPath.Size = new System.Drawing.Size(180, 23);
            this.bGetRebarPath.TabIndex = 21;
            this.bGetRebarPath.Text = "Load Rebar File";
            this.bGetRebarPath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetRebarPath.UseVisualStyleBackColor = true;
            this.bGetRebarPath.Click += new System.EventHandler(this.bGetRebarPath_Click);
            // 
            // tbRebarPath
            // 
            this.tbRebarPath.Enabled = false;
            this.tbRebarPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbRebarPath.Location = new System.Drawing.Point(188, 7);
            this.tbRebarPath.Multiline = true;
            this.tbRebarPath.Name = "tbRebarPath";
            this.tbRebarPath.Size = new System.Drawing.Size(900, 21);
            this.tbRebarPath.TabIndex = 20;
            // 
            // panelBOQ
            // 
            this.panelBOQ.Controls.Add(this.bGetBOQPath);
            this.panelBOQ.Controls.Add(this.tbBOQPath);
            this.panelBOQ.Controls.Add(this.tbFF4Path);
            this.panelBOQ.Controls.Add(this.tbFF3Path);
            this.panelBOQ.Controls.Add(this.bGetFFPath);
            this.panelBOQ.Controls.Add(this.tbFF2Path);
            this.panelBOQ.Controls.Add(this.tbFF1Path);
            this.panelBOQ.Controls.Add(this.tbPlu4Path);
            this.panelBOQ.Controls.Add(this.tbPlu3Path);
            this.panelBOQ.Controls.Add(this.bGetPluPath);
            this.panelBOQ.Controls.Add(this.tbPlu2Path);
            this.panelBOQ.Controls.Add(this.tbPlu1Path);
            this.panelBOQ.Controls.Add(this.tbEle4Path);
            this.panelBOQ.Controls.Add(this.tbEle3Path);
            this.panelBOQ.Controls.Add(this.bGetElePath);
            this.panelBOQ.Controls.Add(this.tbEle2Path);
            this.panelBOQ.Controls.Add(this.tbEle1Path);
            this.panelBOQ.Controls.Add(this.tbMec4Path);
            this.panelBOQ.Controls.Add(this.tbMec3Path);
            this.panelBOQ.Controls.Add(this.bGetMecPath);
            this.panelBOQ.Controls.Add(this.tbMec2Path);
            this.panelBOQ.Controls.Add(this.tbMec1Path);
            this.panelBOQ.Location = new System.Drawing.Point(3, 171);
            this.panelBOQ.Name = "panelBOQ";
            this.panelBOQ.Size = new System.Drawing.Size(1199, 432);
            this.panelBOQ.TabIndex = 8;
            // 
            // bGetBOQPath
            // 
            this.bGetBOQPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetBOQPath.Location = new System.Drawing.Point(5, 6);
            this.bGetBOQPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetBOQPath.Name = "bGetBOQPath";
            this.bGetBOQPath.Size = new System.Drawing.Size(180, 23);
            this.bGetBOQPath.TabIndex = 71;
            this.bGetBOQPath.Text = "Load BOQ File";
            this.bGetBOQPath.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.bGetBOQPath.UseVisualStyleBackColor = true;
            this.bGetBOQPath.Click += new System.EventHandler(this.bGetBOQPath_Click);
            // 
            // tbBOQPath
            // 
            this.tbBOQPath.Enabled = false;
            this.tbBOQPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbBOQPath.Location = new System.Drawing.Point(187, 7);
            this.tbBOQPath.Multiline = true;
            this.tbBOQPath.Name = "tbBOQPath";
            this.tbBOQPath.Size = new System.Drawing.Size(900, 21);
            this.tbBOQPath.TabIndex = 70;
            // 
            // tbFF4Path
            // 
            this.tbFF4Path.Enabled = false;
            this.tbFF4Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbFF4Path.Location = new System.Drawing.Point(188, 367);
            this.tbFF4Path.Multiline = true;
            this.tbFF4Path.Name = "tbFF4Path";
            this.tbFF4Path.Size = new System.Drawing.Size(900, 22);
            this.tbFF4Path.TabIndex = 66;
            // 
            // tbFF3Path
            // 
            this.tbFF3Path.Enabled = false;
            this.tbFF3Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbFF3Path.Location = new System.Drawing.Point(188, 346);
            this.tbFF3Path.Multiline = true;
            this.tbFF3Path.Name = "tbFF3Path";
            this.tbFF3Path.Size = new System.Drawing.Size(900, 22);
            this.tbFF3Path.TabIndex = 65;
            // 
            // bGetFFPath
            // 
            this.bGetFFPath.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetFFPath.Location = new System.Drawing.Point(6, 304);
            this.bGetFFPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetFFPath.Name = "bGetFFPath";
            this.bGetFFPath.Size = new System.Drawing.Size(180, 86);
            this.bGetFFPath.TabIndex = 61;
            this.bGetFFPath.Text = "Load Fire Fighting Files";
            this.bGetFFPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bGetFFPath.UseVisualStyleBackColor = true;
            this.bGetFFPath.Click += new System.EventHandler(this.bGetFF1Path_Click);
            // 
            // tbFF2Path
            // 
            this.tbFF2Path.Enabled = false;
            this.tbFF2Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbFF2Path.Location = new System.Drawing.Point(188, 326);
            this.tbFF2Path.Multiline = true;
            this.tbFF2Path.Name = "tbFF2Path";
            this.tbFF2Path.Size = new System.Drawing.Size(900, 22);
            this.tbFF2Path.TabIndex = 60;
            // 
            // tbFF1Path
            // 
            this.tbFF1Path.Enabled = false;
            this.tbFF1Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbFF1Path.Location = new System.Drawing.Point(188, 305);
            this.tbFF1Path.Multiline = true;
            this.tbFF1Path.Name = "tbFF1Path";
            this.tbFF1Path.Size = new System.Drawing.Size(900, 22);
            this.tbFF1Path.TabIndex = 59;
            // 
            // tbPlu4Path
            // 
            this.tbPlu4Path.Enabled = false;
            this.tbPlu4Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbPlu4Path.Location = new System.Drawing.Point(188, 278);
            this.tbPlu4Path.Multiline = true;
            this.tbPlu4Path.Name = "tbPlu4Path";
            this.tbPlu4Path.Size = new System.Drawing.Size(900, 22);
            this.tbPlu4Path.TabIndex = 53;
            // 
            // tbPlu3Path
            // 
            this.tbPlu3Path.Enabled = false;
            this.tbPlu3Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbPlu3Path.Location = new System.Drawing.Point(188, 257);
            this.tbPlu3Path.Multiline = true;
            this.tbPlu3Path.Name = "tbPlu3Path";
            this.tbPlu3Path.Size = new System.Drawing.Size(900, 22);
            this.tbPlu3Path.TabIndex = 52;
            // 
            // bGetPluPath
            // 
            this.bGetPluPath.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetPluPath.Location = new System.Drawing.Point(6, 215);
            this.bGetPluPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetPluPath.Name = "bGetPluPath";
            this.bGetPluPath.Size = new System.Drawing.Size(180, 86);
            this.bGetPluPath.TabIndex = 48;
            this.bGetPluPath.Text = "Load Plumbing Files";
            this.bGetPluPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bGetPluPath.UseVisualStyleBackColor = true;
            this.bGetPluPath.Click += new System.EventHandler(this.bGetPlu1Path_Click);
            // 
            // tbPlu2Path
            // 
            this.tbPlu2Path.Enabled = false;
            this.tbPlu2Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbPlu2Path.Location = new System.Drawing.Point(188, 237);
            this.tbPlu2Path.Multiline = true;
            this.tbPlu2Path.Name = "tbPlu2Path";
            this.tbPlu2Path.Size = new System.Drawing.Size(900, 22);
            this.tbPlu2Path.TabIndex = 47;
            // 
            // tbPlu1Path
            // 
            this.tbPlu1Path.Enabled = false;
            this.tbPlu1Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbPlu1Path.Location = new System.Drawing.Point(188, 216);
            this.tbPlu1Path.Multiline = true;
            this.tbPlu1Path.Name = "tbPlu1Path";
            this.tbPlu1Path.Size = new System.Drawing.Size(900, 22);
            this.tbPlu1Path.TabIndex = 46;
            // 
            // tbEle4Path
            // 
            this.tbEle4Path.Enabled = false;
            this.tbEle4Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbEle4Path.Location = new System.Drawing.Point(188, 189);
            this.tbEle4Path.Multiline = true;
            this.tbEle4Path.Name = "tbEle4Path";
            this.tbEle4Path.Size = new System.Drawing.Size(900, 22);
            this.tbEle4Path.TabIndex = 40;
            // 
            // tbEle3Path
            // 
            this.tbEle3Path.Enabled = false;
            this.tbEle3Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbEle3Path.Location = new System.Drawing.Point(188, 168);
            this.tbEle3Path.Multiline = true;
            this.tbEle3Path.Name = "tbEle3Path";
            this.tbEle3Path.Size = new System.Drawing.Size(900, 22);
            this.tbEle3Path.TabIndex = 39;
            // 
            // bGetElePath
            // 
            this.bGetElePath.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetElePath.Location = new System.Drawing.Point(6, 126);
            this.bGetElePath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetElePath.Name = "bGetElePath";
            this.bGetElePath.Size = new System.Drawing.Size(180, 86);
            this.bGetElePath.TabIndex = 35;
            this.bGetElePath.Text = "Load Electrical Files";
            this.bGetElePath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bGetElePath.UseVisualStyleBackColor = true;
            this.bGetElePath.Click += new System.EventHandler(this.bGetEle1Path_Click);
            // 
            // tbEle2Path
            // 
            this.tbEle2Path.Enabled = false;
            this.tbEle2Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbEle2Path.Location = new System.Drawing.Point(188, 148);
            this.tbEle2Path.Multiline = true;
            this.tbEle2Path.Name = "tbEle2Path";
            this.tbEle2Path.Size = new System.Drawing.Size(900, 22);
            this.tbEle2Path.TabIndex = 34;
            // 
            // tbEle1Path
            // 
            this.tbEle1Path.Enabled = false;
            this.tbEle1Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbEle1Path.Location = new System.Drawing.Point(188, 127);
            this.tbEle1Path.Multiline = true;
            this.tbEle1Path.Name = "tbEle1Path";
            this.tbEle1Path.Size = new System.Drawing.Size(900, 22);
            this.tbEle1Path.TabIndex = 33;
            // 
            // tbMec4Path
            // 
            this.tbMec4Path.Enabled = false;
            this.tbMec4Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbMec4Path.Location = new System.Drawing.Point(188, 100);
            this.tbMec4Path.Multiline = true;
            this.tbMec4Path.Name = "tbMec4Path";
            this.tbMec4Path.Size = new System.Drawing.Size(900, 22);
            this.tbMec4Path.TabIndex = 27;
            // 
            // tbMec3Path
            // 
            this.tbMec3Path.Enabled = false;
            this.tbMec3Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbMec3Path.Location = new System.Drawing.Point(188, 79);
            this.tbMec3Path.Multiline = true;
            this.tbMec3Path.Name = "tbMec3Path";
            this.tbMec3Path.Size = new System.Drawing.Size(900, 22);
            this.tbMec3Path.TabIndex = 26;
            // 
            // bGetMecPath
            // 
            this.bGetMecPath.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetMecPath.Location = new System.Drawing.Point(6, 37);
            this.bGetMecPath.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.bGetMecPath.Name = "bGetMecPath";
            this.bGetMecPath.Size = new System.Drawing.Size(180, 86);
            this.bGetMecPath.TabIndex = 22;
            this.bGetMecPath.Text = "Load Mechanical Files";
            this.bGetMecPath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bGetMecPath.UseVisualStyleBackColor = true;
            this.bGetMecPath.Click += new System.EventHandler(this.bGetMec1Path_Click);
            // 
            // tbMec2Path
            // 
            this.tbMec2Path.Enabled = false;
            this.tbMec2Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbMec2Path.Location = new System.Drawing.Point(188, 59);
            this.tbMec2Path.Multiline = true;
            this.tbMec2Path.Name = "tbMec2Path";
            this.tbMec2Path.Size = new System.Drawing.Size(900, 22);
            this.tbMec2Path.TabIndex = 21;
            // 
            // tbMec1Path
            // 
            this.tbMec1Path.Enabled = false;
            this.tbMec1Path.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbMec1Path.Location = new System.Drawing.Point(188, 38);
            this.tbMec1Path.Multiline = true;
            this.tbMec1Path.Name = "tbMec1Path";
            this.tbMec1Path.Size = new System.Drawing.Size(900, 22);
            this.tbMec1Path.TabIndex = 20;
            // 
            // bClose
            // 
            this.bClose.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bClose.Location = new System.Drawing.Point(1024, 3);
            this.bClose.Name = "bClose";
            this.bClose.Size = new System.Drawing.Size(61, 26);
            this.bClose.TabIndex = 9;
            this.bClose.Text = "OK";
            this.bClose.UseVisualStyleBackColor = true;
            this.bClose.Click += new System.EventHandler(this.bClose_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Black", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(519, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 22);
            this.label1.TabIndex = 10;
            this.label1.Text = "GET FILE PATHS";
            // 
            // openPath2
            // 
            this.openPath2.Multiselect = true;
            // 
            // bClear
            // 
            this.bClear.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bClear.Location = new System.Drawing.Point(957, 3);
            this.bClear.Name = "bClear";
            this.bClear.Size = new System.Drawing.Size(61, 26);
            this.bClear.TabIndex = 11;
            this.bClear.Text = "Clear";
            this.bClear.UseVisualStyleBackColor = true;
            this.bClear.Click += new System.EventHandler(this.bClear_Click);
            // 
            // FormTTD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1095, 566);
            this.Controls.Add(this.bClear);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bClose);
            this.Controls.Add(this.panelBOQ);
            this.Controls.Add(this.panelRebar);
            this.Controls.Add(this.panelTIO);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormTTD";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Files";
            this.Load += new System.EventHandler(this.FormTTD_Load);
            this.panelTIO.ResumeLayout(false);
            this.panelTIO.PerformLayout();
            this.panelRebar.ResumeLayout(false);
            this.panelRebar.PerformLayout();
            this.panelBOQ.ResumeLayout(false);
            this.panelBOQ.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Panel panelTIO;
        private System.Windows.Forms.Button bGetAExpPath;
        private System.Windows.Forms.Button bGetACatePath;
        private System.Windows.Forms.TextBox tbAExpPath;
        private System.Windows.Forms.TextBox tbACatePath;
        private System.Windows.Forms.Button bGetSExpPath;
        private System.Windows.Forms.Button bGetSCatePath;
        private System.Windows.Forms.TextBox tbSExpPath;
        private System.Windows.Forms.TextBox tbSCatePath;
        private System.Windows.Forms.OpenFileDialog openPath;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panelRebar;
        private System.Windows.Forms.Button bGetRebarPath;
        private System.Windows.Forms.TextBox tbRebarPath;
        private System.Windows.Forms.Panel panelBOQ;
        private System.Windows.Forms.Button bGetMecPath;
        private System.Windows.Forms.TextBox tbMec2Path;
        private System.Windows.Forms.TextBox tbMec1Path;
        private System.Windows.Forms.TextBox tbMec4Path;
        private System.Windows.Forms.TextBox tbMec3Path;
        private System.Windows.Forms.TextBox tbEle4Path;
        private System.Windows.Forms.TextBox tbEle3Path;
        private System.Windows.Forms.Button bGetElePath;
        private System.Windows.Forms.TextBox tbEle2Path;
        private System.Windows.Forms.TextBox tbEle1Path;
        private System.Windows.Forms.Button bGetBOQPath;
        private System.Windows.Forms.TextBox tbBOQPath;
        private System.Windows.Forms.TextBox tbFF4Path;
        private System.Windows.Forms.TextBox tbFF3Path;
        private System.Windows.Forms.Button bGetFFPath;
        private System.Windows.Forms.TextBox tbFF2Path;
        private System.Windows.Forms.TextBox tbFF1Path;
        private System.Windows.Forms.TextBox tbPlu4Path;
        private System.Windows.Forms.TextBox tbPlu3Path;
        private System.Windows.Forms.Button bGetPluPath;
        private System.Windows.Forms.TextBox tbPlu2Path;
        private System.Windows.Forms.TextBox tbPlu1Path;
        private System.Windows.Forms.Button bClose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openPath2;
        private System.Windows.Forms.Button bClear;
    }
}