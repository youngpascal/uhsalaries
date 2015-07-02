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
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.button6 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(125, 134);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(35, 20);
            this.textBox1.TabIndex = 5;
            this.textBox1.Text = "2";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(125, 184);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(35, 20);
            this.textBox3.TabIndex = 7;
            this.textBox3.Text = "1.03";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(125, 210);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(35, 20);
            this.textBox2.TabIndex = 8;
            this.textBox2.Text = "1.015";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Assoc. Constant \'l\':";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 187);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Change constant \'d\':";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 213);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Change constant \'k\':";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(169, 134);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(534, 96);
            this.button1.TabIndex = 10;
            this.button1.Text = "RUN\r\nCALCULATION";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkedFilters
            // 
            this.checkedFilters.CheckOnClick = true;
            this.checkedFilters.FormattingEnabled = true;
            this.checkedFilters.Location = new System.Drawing.Point(211, 134);
            this.checkedFilters.Name = "checkedFilters";
            this.checkedFilters.Size = new System.Drawing.Size(240, 304);
            this.checkedFilters.Sorted = true;
            this.checkedFilters.TabIndex = 13;
            this.checkedFilters.Visible = false;
            // 
            // checkedDepartmentFilters
            // 
            this.checkedDepartmentFilters.CheckOnClick = true;
            this.checkedDepartmentFilters.FormattingEnabled = true;
            this.checkedDepartmentFilters.Location = new System.Drawing.Point(457, 134);
            this.checkedDepartmentFilters.Name = "checkedDepartmentFilters";
            this.checkedDepartmentFilters.Size = new System.Drawing.Size(125, 304);
            this.checkedDepartmentFilters.Sorted = true;
            this.checkedDepartmentFilters.TabIndex = 14;
            this.checkedDepartmentFilters.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 161);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Prof. Constant \'l\':";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(125, 158);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(35, 20);
            this.textBox4.TabIndex = 6;
            this.textBox4.Text = "3";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(15, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(145, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "UH Salary Data";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(15, 40);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(145, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "[1] New Asst. Prof. Salaries";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(15, 70);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(145, 23);
            this.button4.TabIndex = 3;
            this.button4.Text = "[2] Tier 1 Specalty Salaries";
            this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(15, 102);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(145, 23);
            this.button5.TabIndex = 4;
            this.button5.Text = "[3] UH Specalty Weights";
            this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(166, 17);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(115, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "Current path for input 0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(166, 45);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(115, 13);
            this.label6.TabIndex = 22;
            this.label6.Text = "Current path for input 1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(166, 75);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(115, 13);
            this.label7.TabIndex = 23;
            this.label7.Text = "Current path for input 2";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(166, 107);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(115, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "Current path for input 3";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(169, 134);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(193, 44);
            this.button6.TabIndex = 9;
            this.button6.Text = "Get Filters";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(715, 241);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
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
            this.Text = "Salary Statistic Calculator";
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
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button6;

    }
}

