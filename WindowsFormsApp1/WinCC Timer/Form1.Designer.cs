namespace WinCC_Timer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.refIdBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.sqlPathBox = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.button5 = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.checkBox3 = new MyCheckBox();
            this.export = new System.Windows.Forms.Button();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.numberOfMeasurements = new System.Windows.Forms.TextBox();
            this.timeInterval = new System.Windows.Forms.TextBox();
            this.hoursToRun = new System.Windows.Forms.TextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.snapBox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Time!";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(801, 90);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(93, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Manip Grafexe";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // refIdBox
            // 
            this.refIdBox.Location = new System.Drawing.Point(837, 35);
            this.refIdBox.Name = "refIdBox";
            this.refIdBox.Size = new System.Drawing.Size(26, 20);
            this.refIdBox.TabIndex = 4;
            this.refIdBox.Text = "1";
            this.refIdBox.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(798, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "RefId";
            this.label2.Visible = false;
            // 
            // sqlPathBox
            // 
            this.sqlPathBox.Location = new System.Drawing.Point(869, 35);
            this.sqlPathBox.Name = "sqlPathBox";
            this.sqlPathBox.Size = new System.Drawing.Size(74, 20);
            this.sqlPathBox.TabIndex = 6;
            this.sqlPathBox.Text = "TCMHMID01";
            this.sqlPathBox.Visible = false;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(73, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(120, 20);
            this.textBox1.TabIndex = 7;
            this.textBox1.Text = "SDI_CSPM_1-1033.sql";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(279, 39);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(60, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "CalcTimes";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(801, 61);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 9;
            this.button4.Text = "button4";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.Button4_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(11, 267);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(235, 43);
            this.listBox1.TabIndex = 11;
            // 
            // treeView1
            // 
            this.treeView1.Location = new System.Drawing.Point(12, 70);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(361, 191);
            this.treeView1.TabIndex = 12;
            this.treeView1.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterCheck);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(317, 267);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(62, 17);
            this.checkBox1.TabIndex = 13;
            this.checkBox1.Text = "Expand";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(317, 293);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(56, 17);
            this.checkBox2.TabIndex = 14;
            this.checkBox2.Text = "Select";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(199, 38);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 15;
            this.button5.Text = "StdDev";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(73, 39);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(120, 20);
            this.textBox3.TabIndex = 16;
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(398, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(363, 298);
            this.listView1.TabIndex = 17;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.Visible = false;
            // 
            // checkBox3
            // 
            this.checkBox3.Location = new System.Drawing.Point(327, 10);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Padding = new System.Windows.Forms.Padding(6);
            this.checkBox3.Size = new System.Drawing.Size(46, 24);
            this.checkBox3.TabIndex = 0;
            this.checkBox3.CheckedChanged += new System.EventHandler(this.checkBox3_CheckedChanged);
            // 
            // export
            // 
            this.export.Location = new System.Drawing.Point(279, 12);
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(46, 23);
            this.export.TabIndex = 18;
            this.export.Text = "Export";
            this.export.UseVisualStyleBackColor = true;
            this.export.Click += new System.EventHandler(this.button6_Click);
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(12, 41);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(60, 17);
            this.checkBox4.TabIndex = 19;
            this.checkBox4.Text = "Screen";
            this.checkBox4.UseVisualStyleBackColor = true;
            this.checkBox4.CheckedChanged += new System.EventHandler(this.checkBox4_CheckedChanged);
            // 
            // numberOfMeasurements
            // 
            this.numberOfMeasurements.Location = new System.Drawing.Point(226, 12);
            this.numberOfMeasurements.Name = "numberOfMeasurements";
            this.numberOfMeasurements.Size = new System.Drawing.Size(20, 20);
            this.numberOfMeasurements.TabIndex = 20;
            this.numberOfMeasurements.Text = "16";
            // 
            // timeInterval
            // 
            this.timeInterval.Location = new System.Drawing.Point(253, 12);
            this.timeInterval.Name = "timeInterval";
            this.timeInterval.Size = new System.Drawing.Size(20, 20);
            this.timeInterval.TabIndex = 21;
            this.timeInterval.Text = "30";
            // 
            // hoursToRun
            // 
            this.hoursToRun.Location = new System.Drawing.Point(199, 12);
            this.hoursToRun.Name = "hoursToRun";
            this.hoursToRun.Size = new System.Drawing.Size(20, 20);
            this.hoursToRun.TabIndex = 22;
            this.hoursToRun.Text = "8";
            this.hoursToRun.TextChanged += new System.EventHandler(this.hoursToRun_TextChanged);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(345, 38);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(28, 23);
            this.button6.TabIndex = 23;
            this.button6.Text = "XL";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click_1);
            // 
            // snapBox
            // 
            this.snapBox.AutoSize = true;
            this.snapBox.Location = new System.Drawing.Point(253, 267);
            this.snapBox.Name = "snapBox";
            this.snapBox.Size = new System.Drawing.Size(51, 17);
            this.snapBox.TabIndex = 24;
            this.snapBox.Text = "Snap";
            this.snapBox.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(386, 324);
            this.Controls.Add(this.snapBox);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.hoursToRun);
            this.Controls.Add(this.timeInterval);
            this.Controls.Add(this.numberOfMeasurements);
            this.Controls.Add(this.checkBox4);
            this.Controls.Add(this.export);
            this.Controls.Add(this.checkBox3);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.sqlPathBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.refIdBox);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "WinCC Timer";
            this.TopMost = true;
            this.DoubleClick += new System.EventHandler(this.Form1_DoubleClicked);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox refIdBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox sqlPathBox;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.ListView listView1;
        private MyCheckBox checkBox3;
        private System.Windows.Forms.Button export;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.TextBox numberOfMeasurements;
        private System.Windows.Forms.TextBox timeInterval;
        private System.Windows.Forms.TextBox hoursToRun;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.CheckBox snapBox;
    }
}

