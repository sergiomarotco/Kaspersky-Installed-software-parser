namespace Kaspersky_Installed_software_parser
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.BoxApplications = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.BoxWhite = new System.Windows.Forms.TextBox();
            this.BoxBad = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.BadRichTextBox = new System.Windows.Forms.RichTextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.WhiteRichTextBox = new System.Windows.Forms.RichTextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Software = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DNSname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(177, 8);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select file";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // BoxApplications
            // 
            this.BoxApplications.Location = new System.Drawing.Point(13, 60);
            this.BoxApplications.Name = "BoxApplications";
            this.BoxApplications.Size = new System.Drawing.Size(351, 22);
            this.BoxApplications.TabIndex = 1;
            this.BoxApplications.Text = "installed software.xls";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(161, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Selected file for parsing:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 91);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(127, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "White applications:";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(177, 88);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(103, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "Select file";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // BoxWhite
            // 
            this.BoxWhite.Location = new System.Drawing.Point(13, 117);
            this.BoxWhite.Name = "BoxWhite";
            this.BoxWhite.Size = new System.Drawing.Size(351, 22);
            this.BoxWhite.TabIndex = 13;
            this.BoxWhite.Text = "Applications_white.txt";
            // 
            // BoxBad
            // 
            this.BoxBad.Location = new System.Drawing.Point(13, 174);
            this.BoxBad.Name = "BoxBad";
            this.BoxBad.Size = new System.Drawing.Size(351, 22);
            this.BoxBad.TabIndex = 16;
            this.BoxBad.Text = "Applications_bad.txt";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 148);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(116, 17);
            this.label6.TabIndex = 15;
            this.label6.Text = "Bad applications:";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(177, 145);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(103, 23);
            this.button3.TabIndex = 14;
            this.button3.Text = "Select file";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(280, 11);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(0, 17);
            this.label7.TabIndex = 17;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(13, 202);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 21;
            this.button4.Text = "Export";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.Button4_Click);
            // 
            // BadRichTextBox
            // 
            this.BadRichTextBox.Location = new System.Drawing.Point(13, 248);
            this.BadRichTextBox.Name = "BadRichTextBox";
            this.BadRichTextBox.Size = new System.Drawing.Size(351, 230);
            this.BadRichTextBox.TabIndex = 22;
            this.BadRichTextBox.Text = "";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(94, 202);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(134, 23);
            this.button5.TabIndex = 23;
            this.button5.Text = "Add to Bad";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.ButtonBadADD_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 228);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 17);
            this.label2.TabIndex = 24;
            this.label2.Text = "New bad software:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 481);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(132, 17);
            this.label3.TabIndex = 26;
            this.label3.Text = "New white software:";
            // 
            // WhiteRichTextBox
            // 
            this.WhiteRichTextBox.Location = new System.Drawing.Point(13, 501);
            this.WhiteRichTextBox.Name = "WhiteRichTextBox";
            this.WhiteRichTextBox.Size = new System.Drawing.Size(351, 253);
            this.WhiteRichTextBox.TabIndex = 25;
            this.WhiteRichTextBox.Text = "";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(234, 202);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(130, 23);
            this.button6.TabIndex = 27;
            this.button6.Text = "Add to White";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.ButtonWhiteADD_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(13, 760);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(178, 23);
            this.button7.TabIndex = 28;
            this.button7.Text = "Сохранить списки";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.Button7_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(10, 786);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(175, 17);
            this.linkLabel1.TabIndex = 29;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "sergiomarotco@gmail.com";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel1_LinkClicked);
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(229, 786);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(135, 17);
            this.linkLabel2.TabIndex = 30;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "icons by icons8.com";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel2_LinkClicked);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.checkBox3);
            this.splitContainer1.Panel1.Controls.Add(this.checkBox2);
            this.splitContainer1.Panel1.Controls.Add(this.checkBox1);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.linkLabel2);
            this.splitContainer1.Panel1.Controls.Add(this.BadRichTextBox);
            this.splitContainer1.Panel1.Controls.Add(this.linkLabel1);
            this.splitContainer1.Panel1.Controls.Add(this.button1);
            this.splitContainer1.Panel1.Controls.Add(this.button7);
            this.splitContainer1.Panel1.Controls.Add(this.BoxApplications);
            this.splitContainer1.Panel1.Controls.Add(this.button6);
            this.splitContainer1.Panel1.Controls.Add(this.button2);
            this.splitContainer1.Panel1.Controls.Add(this.label3);
            this.splitContainer1.Panel1.Controls.Add(this.label5);
            this.splitContainer1.Panel1.Controls.Add(this.WhiteRichTextBox);
            this.splitContainer1.Panel1.Controls.Add(this.BoxWhite);
            this.splitContainer1.Panel1.Controls.Add(this.label2);
            this.splitContainer1.Panel1.Controls.Add(this.button3);
            this.splitContainer1.Panel1.Controls.Add(this.button5);
            this.splitContainer1.Panel1.Controls.Add(this.label6);
            this.splitContainer1.Panel1.Controls.Add(this.button4);
            this.splitContainer1.Panel1.Controls.Add(this.BoxBad);
            this.splitContainer1.Panel1.Controls.Add(this.label7);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridView1);
            this.splitContainer1.Size = new System.Drawing.Size(1571, 807);
            this.splitContainer1.SplitterDistance = 368;
            this.splitContainer1.TabIndex = 31;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(13, 33);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(77, 21);
            this.checkBox1.TabIndex = 31;
            this.checkBox1.Text = "Lookup";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Software,
            this.IP,
            this.DNSname,
            this.Type});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1199, 807);
            this.dataGridView1.TabIndex = 21;
            this.dataGridView1.DoubleClick += new System.EventHandler(this.DataGridView1_DoubleClick);
            // 
            // Software
            // 
            this.Software.HeaderText = "Software";
            this.Software.Name = "Software";
            this.Software.Width = 92;
            // 
            // IP
            // 
            this.IP.HeaderText = "IP";
            this.IP.Name = "IP";
            this.IP.Width = 49;
            // 
            // DNSname
            // 
            this.DNSname.HeaderText = "DNS name";
            this.DNSname.Name = "DNSname";
            this.DNSname.Width = 105;
            // 
            // Type
            // 
            this.Type.HeaderText = "Type";
            this.Type.Name = "Type";
            this.Type.Width = 69;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog1";
            this.openFileDialog2.Filter = "TXT files (*.txt)|*.txt";
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(94, 33);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(100, 21);
            this.checkBox2.TabIndex = 32;
            this.checkBox2.Text = "Show white";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Checked = true;
            this.checkBox3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox3.Location = new System.Drawing.Point(200, 33);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(155, 21);
            this.checkBox3.TabIndex = 33;
            this.checkBox3.Text = "Only \'Need request\'";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1571, 807);
            this.Controls.Add(this.splitContainer1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Kaspersky installed software parser";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox BoxApplications;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox BoxWhite;
        private System.Windows.Forms.TextBox BoxBad;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.RichTextBox BadRichTextBox;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox WhiteRichTextBox;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Software;
        private System.Windows.Forms.DataGridViewTextBoxColumn IP;
        private System.Windows.Forms.DataGridViewTextBoxColumn DNSname;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox3;
    }
}

