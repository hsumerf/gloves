namespace Accounts
{
    partial class AddDeal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddDeal));
            this.label2 = new System.Windows.Forms.Label();
            this.companyBox = new System.Windows.Forms.ComboBox();
            this.itemBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.nameBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.labelitems = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.valQty = new System.Windows.Forms.NumericUpDown();
            this.ValPer = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.total = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.valQty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ValPer)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "Company";
            // 
            // companyBox
            // 
            this.companyBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.companyBox.FormattingEnabled = true;
            this.companyBox.Location = new System.Drawing.Point(84, 33);
            this.companyBox.Name = "companyBox";
            this.companyBox.Size = new System.Drawing.Size(105, 23);
            this.companyBox.Sorted = true;
            this.companyBox.TabIndex = 4;
            this.companyBox.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            this.companyBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.companyBox_KeyPress);
            // 
            // itemBox
            // 
            this.itemBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.itemBox.FormattingEnabled = true;
            this.itemBox.Location = new System.Drawing.Point(233, 33);
            this.itemBox.Name = "itemBox";
            this.itemBox.Size = new System.Drawing.Size(157, 23);
            this.itemBox.Sorted = true;
            this.itemBox.TabIndex = 5;
            this.itemBox.DropDown += new System.EventHandler(this.itemBox_DropDown);
            this.itemBox.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            this.itemBox.Enter += new System.EventHandler(this.itemBox_Enter);
            this.itemBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.itemBox_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(196, 37);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 15);
            this.label3.TabIndex = 5;
            this.label3.Text = "Item";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.companyBox);
            this.groupBox1.Controls.Add(this.itemBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 119);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(399, 87);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Product";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.nameBox);
            this.groupBox2.Location = new System.Drawing.Point(10, 21);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(400, 87);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Client";
            this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 38);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 15);
            this.label4.TabIndex = 1;
            this.label4.Text = "Name";
            // 
            // nameBox
            // 
            this.nameBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.nameBox.FormattingEnabled = true;
            this.nameBox.Location = new System.Drawing.Point(65, 35);
            this.nameBox.Name = "nameBox";
            this.nameBox.Size = new System.Drawing.Size(313, 23);
            this.nameBox.Sorted = true;
            this.nameBox.TabIndex = 3;
            this.nameBox.SelectedIndexChanged += new System.EventHandler(this.nameBox_SelectedIndexChanged);
            this.nameBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.nameBox_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 105);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 15);
            this.label1.TabIndex = 8;
            this.label1.Text = "Date";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(65, 102);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(167, 21);
            this.dateTimePicker1.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.label5.Location = new System.Drawing.Point(17, 69);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(24, 15);
            this.label5.TabIndex = 10;
            this.label5.Text = "Qty";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(227, 367);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(87, 27);
            this.button2.TabIndex = 16;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // labelitems
            // 
            this.labelitems.AutoSize = true;
            this.labelitems.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelitems.Location = new System.Drawing.Point(146, 72);
            this.labelitems.Name = "labelitems";
            this.labelitems.Size = new System.Drawing.Size(0, 13);
            this.labelitems.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(136, 70);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(10, 15);
            this.label7.TabIndex = 17;
            this.label7.Text = "/";
            // 
            // valQty
            // 
            this.valQty.Location = new System.Drawing.Point(63, 66);
            this.valQty.Maximum = new decimal(new int[] {
            999999,
            0,
            0,
            0});
            this.valQty.Name = "valQty";
            this.valQty.Size = new System.Drawing.Size(70, 21);
            this.valQty.TabIndex = 7;
            this.valQty.ValueChanged += new System.EventHandler(this.valQty_ValueChanged);
            this.valQty.KeyDown += new System.Windows.Forms.KeyEventHandler(this.valQty_KeyDown);
            // 
            // ValPer
            // 
            this.ValPer.DecimalPlaces = 2;
            this.ValPer.Location = new System.Drawing.Point(274, 66);
            this.ValPer.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.ValPer.Name = "ValPer";
            this.ValPer.Size = new System.Drawing.Size(90, 21);
            this.ValPer.TabIndex = 8;
            this.ValPer.ValueChanged += new System.EventHandler(this.ValPer_ValueChanged);
            this.ValPer.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ValPer_KeyDown);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.label6.Location = new System.Drawing.Point(234, 70);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 15);
            this.label6.TabIndex = 19;
            this.label6.Text = "Rate";
            // 
            // total
            // 
            this.total.AutoSize = true;
            this.total.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.total.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.total.Location = new System.Drawing.Point(252, 103);
            this.total.Name = "total";
            this.total.Size = new System.Drawing.Size(52, 16);
            this.total.TabIndex = 21;
            this.total.Text = "Total: ";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.labelitems);
            this.groupBox3.Controls.Add(this.total);
            this.groupBox3.Controls.Add(this.ValPer);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.valQty);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.dateTimePicker1);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(10, 212);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(401, 143);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Particulars";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(65, 28);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(313, 21);
            this.textBox1.TabIndex = 6;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 32);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 15);
            this.label8.TabIndex = 22;
            this.label8.Text = "Details";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(324, 367);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(87, 27);
            this.button1.TabIndex = 10;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // AddDeal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(423, 406);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddDeal";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sell ";
            this.Load += new System.EventHandler(this.AddDeal_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.valQty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ValPer)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox companyBox;
        private System.Windows.Forms.ComboBox itemBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox nameBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label labelitems;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown valQty;
        private System.Windows.Forms.NumericUpDown ValPer;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label total;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label8;
    }
}