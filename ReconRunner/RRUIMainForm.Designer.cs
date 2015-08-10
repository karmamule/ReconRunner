namespace ReconRunner
{
    partial class RRUIMainForm
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
            this.fileSaveDialog = new System.Windows.Forms.SaveFileDialog();
            this.statusText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.readFromXMLFile = new System.Windows.Forms.Button();
            this.xmlFileOpenDialog = new System.Windows.Forms.OpenFileDialog();
            this.createReconXMLFile = new System.Windows.Forms.Button();
            this.runRecons = new System.Windows.Forms.Button();
            this.createSampleReconXMLFile = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // statusText
            // 
            this.statusText.Location = new System.Drawing.Point(27, 115);
            this.statusText.Multiline = true;
            this.statusText.Name = "statusText";
            this.statusText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.statusText.Size = new System.Drawing.Size(388, 302);
            this.statusText.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Status:";
            // 
            // readFromXMLFile
            // 
            this.readFromXMLFile.Location = new System.Drawing.Point(27, 26);
            this.readFromXMLFile.Name = "readFromXMLFile";
            this.readFromXMLFile.Size = new System.Drawing.Size(153, 23);
            this.readFromXMLFile.TabIndex = 3;
            this.readFromXMLFile.Text = "Load Sources File";
            this.readFromXMLFile.UseVisualStyleBackColor = true;
            this.readFromXMLFile.Click += new System.EventHandler(this.readFromSourcesXMLFile_Click);
            // 
            // createReconXMLFile
            // 
            this.createReconXMLFile.Enabled = false;
            this.createReconXMLFile.Location = new System.Drawing.Point(27, 423);
            this.createReconXMLFile.Name = "createReconXMLFile";
            this.createReconXMLFile.Size = new System.Drawing.Size(153, 23);
            this.createReconXMLFile.TabIndex = 0;
            this.createReconXMLFile.Text = "Write Recon XML File";
            this.createReconXMLFile.UseVisualStyleBackColor = true;
            this.createReconXMLFile.Visible = false;
            this.createReconXMLFile.Click += new System.EventHandler(this.createReconXMLFile_Click);
            // 
            // runRecons
            // 
            this.runRecons.Enabled = false;
            this.runRecons.Location = new System.Drawing.Point(225, 39);
            this.runRecons.Name = "runRecons";
            this.runRecons.Size = new System.Drawing.Size(151, 23);
            this.runRecons.TabIndex = 4;
            this.runRecons.Text = "Run Recons";
            this.runRecons.UseVisualStyleBackColor = true;
            this.runRecons.Click += new System.EventHandler(this.runRecons_Click);
            // 
            // createSampleReconXMLFile
            // 
            this.createSampleReconXMLFile.Location = new System.Drawing.Point(264, 423);
            this.createSampleReconXMLFile.Name = "createSampleReconXMLFile";
            this.createSampleReconXMLFile.Size = new System.Drawing.Size(151, 23);
            this.createSampleReconXMLFile.TabIndex = 5;
            this.createSampleReconXMLFile.Text = "Create Sample Recon Files";
            this.createSampleReconXMLFile.UseVisualStyleBackColor = true;
            this.createSampleReconXMLFile.Click += new System.EventHandler(this.createSampleReconXMLFile_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(27, 55);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(153, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Load Recons File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.readFromReconsXMLFile_Click);
            // 
            // RRUIMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 458);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.createSampleReconXMLFile);
            this.Controls.Add(this.runRecons);
            this.Controls.Add(this.readFromXMLFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.statusText);
            this.Controls.Add(this.createReconXMLFile);
            this.Name = "RRUIMainForm";
            this.Text = "Recon Runner";
            this.Load += new System.EventHandler(this.RRUIMainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.SaveFileDialog fileSaveDialog;
        private System.Windows.Forms.TextBox statusText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button readFromXMLFile;
        private System.Windows.Forms.OpenFileDialog xmlFileOpenDialog;
        private System.Windows.Forms.Button createReconXMLFile;
        private System.Windows.Forms.Button runRecons;
        private System.Windows.Forms.Button createSampleReconXMLFile;
        private System.Windows.Forms.Button button1;
    }
}

