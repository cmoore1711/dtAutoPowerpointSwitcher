namespace dtPowerpointUpdater
{
    partial class Form1
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
            this.updateFileLocationBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.fileLocationTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // updateFileLocationBtn
            // 
            this.updateFileLocationBtn.Font = new System.Drawing.Font("Segoe UI Semibold", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.updateFileLocationBtn.Location = new System.Drawing.Point(385, 123);
            this.updateFileLocationBtn.Name = "updateFileLocationBtn";
            this.updateFileLocationBtn.Size = new System.Drawing.Size(163, 35);
            this.updateFileLocationBtn.TabIndex = 0;
            this.updateFileLocationBtn.Text = "Change file being monitored";
            this.updateFileLocationBtn.UseVisualStyleBackColor = true;
            this.updateFileLocationBtn.Click += new System.EventHandler(this.updateFileLocationBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(350, 32);
            this.label1.TabIndex = 1;
            this.label1.Text = "DT Powerpoint Auto-Updater";
            // 
            // fileLocationTxt
            // 
            this.fileLocationTxt.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileLocationTxt.Location = new System.Drawing.Point(13, 126);
            this.fileLocationTxt.Name = "fileLocationTxt";
            this.fileLocationTxt.ReadOnly = true;
            this.fileLocationTxt.Size = new System.Drawing.Size(366, 29);
            this.fileLocationTxt.TabIndex = 2;
            this.fileLocationTxt.Text = "C:\\Users\\cmoor\\Desktop\\PowerpointTest";
            this.fileLocationTxt.TextChanged += new System.EventHandler(this.fileLocationTxt_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Semibold", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(146, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "File to Monitor:";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Segoe UI Semibold", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(9, 247);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(286, 66);
            this.label3.TabIndex = 4;
            this.label3.Text = "This program will close and re-open the powerpoint file at the given path. Make s" +
    "ure it is a .pptx file!";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(560, 322);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.fileLocationTxt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.updateFileLocationBtn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button updateFileLocationBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox fileLocationTxt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Timer timer1;
    }
}

