namespace DAPHT
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
            this.buttonCreate = new System.Windows.Forms.Button();
            this.embedfile = new System.Windows.Forms.TextBox();
            this.checkBoxWord = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.iconfile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxOutputLocation = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.mytext = new System.Windows.Forms.RichTextBox();
            this.password = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.checkBoxPowerpoint = new System.Windows.Forms.CheckBox();
            this.checkBoxExcel = new System.Windows.Forms.CheckBox();
            this.checkBoxZip = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.embedfile2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.embedfile3 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonCreate
            // 
            this.buttonCreate.Location = new System.Drawing.Point(350, 43);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(86, 50);
            this.buttonCreate.TabIndex = 0;
            this.buttonCreate.Text = "Create";
            this.buttonCreate.UseVisualStyleBackColor = true;
            this.buttonCreate.Click += new System.EventHandler(this.button1_Click);
            // 
            // embedfile
            // 
            this.embedfile.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.embedfile.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories;
            this.embedfile.Location = new System.Drawing.Point(29, 25);
            this.embedfile.Name = "embedfile";
            this.embedfile.Size = new System.Drawing.Size(292, 20);
            this.embedfile.TabIndex = 1;
            this.embedfile.Text = "C:\\files\\shell.exe";

            // 
            // checkBoxWord
            // 
            this.checkBoxWord.AutoSize = true;
            this.checkBoxWord.Location = new System.Drawing.Point(32, 415);
            this.checkBoxWord.Name = "checkBoxWord";
            this.checkBoxWord.Size = new System.Drawing.Size(52, 17);
            this.checkBoxWord.TabIndex = 2;
            this.checkBoxWord.Text = "Word";
            this.checkBoxWord.UseVisualStyleBackColor = true;
            
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Payload 1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 153);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Use icon from:";
            // 
            // iconfile
            // 
            this.iconfile.Location = new System.Drawing.Point(32, 175);
            this.iconfile.Name = "iconfile";
            this.iconfile.Size = new System.Drawing.Size(292, 20);
            this.iconfile.TabIndex = 5;
            this.iconfile.Text = "C:\\files\\whatever.exe";
            
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(29, 251);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(165, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Required message for documents";
            // 
            // textBoxOutputLocation
            // 
            this.textBoxOutputLocation.Location = new System.Drawing.Point(32, 223);
            this.textBoxOutputLocation.Name = "textBoxOutputLocation";
            this.textBoxOutputLocation.Size = new System.Drawing.Size(292, 20);
            this.textBoxOutputLocation.TabIndex = 9;
            this.textBoxOutputLocation.Text = "C:\\files\\payloads\\";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 207);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Output location:";
            // 
            // mytext
            // 
            this.mytext.Location = new System.Drawing.Point(32, 268);
            this.mytext.Name = "mytext";
            this.mytext.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.mytext.Size = new System.Drawing.Size(344, 80);
            this.mytext.TabIndex = 10;
            this.mytext.Text = "<insert here>";
            // 
            // password
            // 
            this.password.Location = new System.Drawing.Point(32, 367);
            this.password.Name = "password";
            this.password.Size = new System.Drawing.Size(292, 20);
            this.password.TabIndex = 12;
            this.password.Text = "password";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(32, 351);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(132, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Password (where required)";
            // 
            // checkBoxPowerpoint
            // 
            this.checkBoxPowerpoint.AutoSize = true;
            this.checkBoxPowerpoint.Location = new System.Drawing.Point(90, 415);
            this.checkBoxPowerpoint.Name = "checkBoxPowerpoint";
            this.checkBoxPowerpoint.Size = new System.Drawing.Size(80, 17);
            this.checkBoxPowerpoint.TabIndex = 13;
            this.checkBoxPowerpoint.Text = "PowerPoint";
            this.checkBoxPowerpoint.UseVisualStyleBackColor = true;
            
            // 
            // checkBoxExcel
            // 
            this.checkBoxExcel.AutoSize = true;
            this.checkBoxExcel.Location = new System.Drawing.Point(175, 415);
            this.checkBoxExcel.Name = "checkBoxExcel";
            this.checkBoxExcel.Size = new System.Drawing.Size(52, 17);
            this.checkBoxExcel.TabIndex = 14;
            this.checkBoxExcel.Text = "Excel";
            this.checkBoxExcel.UseVisualStyleBackColor = true;
            // 
            // checkBoxZip
            // 
            this.checkBoxZip.AutoSize = true;
            this.checkBoxZip.Checked = true;
            this.checkBoxZip.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxZip.Location = new System.Drawing.Point(233, 415);
            this.checkBoxZip.Name = "checkBoxZip";
            this.checkBoxZip.Size = new System.Drawing.Size(41, 17);
            this.checkBoxZip.TabIndex = 15;
            this.checkBoxZip.Text = "Zip";
            this.checkBoxZip.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBoxZip.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(29, 87);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Payload 3";
            // 
            // embedfile2
            // 
            this.embedfile2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.embedfile2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories;
            this.embedfile2.Location = new System.Drawing.Point(29, 64);
            this.embedfile2.Name = "embedfile2";
            this.embedfile2.Size = new System.Drawing.Size(292, 20);
            this.embedfile2.TabIndex = 16;
            this.embedfile2.Text = "C:\\files\\shell.bat";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(29, 48);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(54, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Payload 2";
            // 
            // embedfile3
            // 
            this.embedfile3.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.embedfile3.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories;
            this.embedfile3.Location = new System.Drawing.Point(29, 103);
            this.embedfile3.Name = "embedfile3";
            this.embedfile3.Size = new System.Drawing.Size(292, 20);
            this.embedfile3.TabIndex = 18;
            this.embedfile3.Text = "C:\\files\\shell.vbs";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 459);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.embedfile3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.embedfile2);
            this.Controls.Add(this.checkBoxZip);
            this.Controls.Add(this.checkBoxExcel);
            this.Controls.Add(this.checkBoxPowerpoint);
            this.Controls.Add(this.password);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.mytext);
            this.Controls.Add(this.textBoxOutputLocation);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.iconfile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkBoxWord);
            this.Controls.Add(this.embedfile);
            this.Controls.Add(this.buttonCreate);
            this.Name = "Form1";
            this.RightToLeftLayout = true;
            this.Text = "Document and Archive Payload Hiding Tool (DAPHT) Beta 0.1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox embedfile;
        private System.Windows.Forms.TextBox iconfile;
        private System.Windows.Forms.TextBox textBoxOutputLocation;
        private System.Windows.Forms.Button buttonCreate;
        private System.Windows.Forms.CheckBox checkBoxWord;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox mytext;
        private System.Windows.Forms.TextBox password;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox checkBoxPowerpoint;
        private System.Windows.Forms.CheckBox checkBoxExcel;
        private System.Windows.Forms.CheckBox checkBoxZip;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox embedfile2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox embedfile3;
    }
}

