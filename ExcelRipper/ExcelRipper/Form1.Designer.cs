namespace ExcelRipper
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
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tb_files = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_range = new System.Windows.Forms.TextBox();
            this.loadButton = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tb_outfile = new System.Windows.Forms.TextBox();
            this.saveButton = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.tb_columnheader = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.DoWorkSonButton = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Target Files";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "(*.xls;*.xlsx;)|*.xls;*.xlsx;|All Files (*.*)|*";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // tb_files
            // 
            this.tb_files.Location = new System.Drawing.Point(15, 26);
            this.tb_files.Multiline = true;
            this.tb_files.Name = "tb_files";
            this.tb_files.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_files.Size = new System.Drawing.Size(400, 95);
            this.tb_files.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 128);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Cell Range";
            // 
            // tb_range
            // 
            this.tb_range.Location = new System.Drawing.Point(15, 145);
            this.tb_range.Name = "tb_range";
            this.tb_range.Size = new System.Drawing.Size(400, 20);
            this.tb_range.TabIndex = 3;
            // 
            // loadButton
            // 
            this.loadButton.Location = new System.Drawing.Point(421, 26);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(75, 23);
            this.loadButton.TabIndex = 4;
            this.loadButton.Text = "Select";
            this.loadButton.UseVisualStyleBackColor = true;
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(420, 148);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "E.x.: C5:C10, C4, D5";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "OutFile:";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "(*.xls;)|*.xls;|All Files (*.*)|*";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // tb_outfile
            // 
            this.tb_outfile.Location = new System.Drawing.Point(3, 25);
            this.tb_outfile.Name = "tb_outfile";
            this.tb_outfile.Size = new System.Drawing.Size(400, 20);
            this.tb_outfile.TabIndex = 7;
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(421, 193);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(75, 23);
            this.saveButton.TabIndex = 8;
            this.saveButton.Text = "Select";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 53);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Column Header";
            // 
            // tb_columnheader
            // 
            this.tb_columnheader.Location = new System.Drawing.Point(3, 69);
            this.tb_columnheader.Name = "tb_columnheader";
            this.tb_columnheader.Size = new System.Drawing.Size(400, 20);
            this.tb_columnheader.TabIndex = 10;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.tb_columnheader);
            this.panel1.Controls.Add(this.tb_outfile);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(12, 171);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(403, 96);
            this.panel1.TabIndex = 11;
            // 
            // DoWorkSonButton
            // 
            this.DoWorkSonButton.Location = new System.Drawing.Point(421, 278);
            this.DoWorkSonButton.Name = "DoWorkSonButton";
            this.DoWorkSonButton.Size = new System.Drawing.Size(75, 23);
            this.DoWorkSonButton.TabIndex = 12;
            this.DoWorkSonButton.Text = "Rip it!";
            this.DoWorkSonButton.UseVisualStyleBackColor = true;
            this.DoWorkSonButton.Click += new System.EventHandler(this.DoWorkSonButton_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 305);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(529, 22);
            this.statusStrip1.TabIndex = 13;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 327);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.DoWorkSonButton);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.loadButton);
            this.Controls.Add(this.tb_range);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tb_files);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "ExcelRipper v2 by EvonsDesigns";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox tb_files;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_range;
        private System.Windows.Forms.Button loadButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TextBox tb_outfile;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tb_columnheader;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button DoWorkSonButton;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
    }
}

