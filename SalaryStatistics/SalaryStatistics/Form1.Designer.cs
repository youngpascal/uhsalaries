namespace SalaryStatistics
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.checkedFilters = new System.Windows.Forms.CheckedListBox();
            this.checkedDepartmentFilters = new System.Windows.Forms.CheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(125, 0);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(35, 20);
            this.textBox1.TabIndex = 2;
            this.textBox1.Text = "2";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(125, 60);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(35, 20);
            this.textBox3.TabIndex = 4;
            this.textBox3.Text = "1.03";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(125, 86);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(35, 20);
            this.textBox2.TabIndex = 6;
            this.textBox2.Text = "1.015";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Assoc. Constant \'l\':";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Change constant \'d\':";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Change constant \'k\':";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(269, 313);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(137, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "Run!";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkedFilters
            // 
            this.checkedFilters.CheckOnClick = true;
            this.checkedFilters.FormattingEnabled = true;
            this.checkedFilters.Location = new System.Drawing.Point(176, 3);
            this.checkedFilters.Name = "checkedFilters";
            this.checkedFilters.Size = new System.Drawing.Size(230, 304);
            this.checkedFilters.Sorted = true;
            this.checkedFilters.TabIndex = 13;
            // 
            // checkedDepartmentFilters
            // 
            this.checkedDepartmentFilters.CheckOnClick = true;
            this.checkedDepartmentFilters.FormattingEnabled = true;
            this.checkedDepartmentFilters.Location = new System.Drawing.Point(412, 3);
            this.checkedDepartmentFilters.Name = "checkedDepartmentFilters";
            this.checkedDepartmentFilters.Size = new System.Drawing.Size(125, 304);
            this.checkedDepartmentFilters.Sorted = true;
            this.checkedDepartmentFilters.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Prof. Constant \'l\':";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(125, 29);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(35, 20);
            this.textBox4.TabIndex = 15;
            this.textBox4.Text = "3";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(548, 346);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.checkedDepartmentFilters);
            this.Controls.Add(this.checkedFilters);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Salary Statistics";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckedListBox checkedFilters;
        private System.Windows.Forms.CheckedListBox checkedDepartmentFilters;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox4;

    }
}

