namespace Report
{
    partial class MainForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
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
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbxName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbxLabel = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtPicker = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.chbxMale = new System.Windows.Forms.CheckBox();
            this.chbxFemale = new System.Windows.Forms.CheckBox();
            this.chlbxAnalysis = new System.Windows.Forms.CheckedListBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.excelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbxName
            // 
            this.tbxName.Location = new System.Drawing.Point(181, 10);
            this.tbxName.Name = "tbxName";
            this.tbxName.Size = new System.Drawing.Size(148, 20);
            this.tbxName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Введите ФИО";
            // 
            // tbxLabel
            // 
            this.tbxLabel.Location = new System.Drawing.Point(181, 108);
            this.tbxLabel.Name = "tbxLabel";
            this.tbxLabel.Size = new System.Drawing.Size(148, 20);
            this.tbxLabel.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Введите название отчета";
            // 
            // dtPicker
            // 
            this.dtPicker.Location = new System.Drawing.Point(181, 157);
            this.dtPicker.Name = "dtPicker";
            this.dtPicker.Size = new System.Drawing.Size(148, 20);
            this.dtPicker.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 163);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Выберите дату";
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(29, 216);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 36);
            this.okButton.TabIndex = 8;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(26, 69);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Ваш пол";
            // 
            // chbxMale
            // 
            this.chbxMale.AutoSize = true;
            this.chbxMale.Location = new System.Drawing.Point(201, 46);
            this.chbxMale.Name = "chbxMale";
            this.chbxMale.Size = new System.Drawing.Size(72, 17);
            this.chbxMale.TabIndex = 10;
            this.chbxMale.Text = "Мужской";
            this.chbxMale.UseVisualStyleBackColor = true;
            this.chbxMale.CheckedChanged += new System.EventHandler(this.chbxMale_CheckedChanged);
            // 
            // chbxFemale
            // 
            this.chbxFemale.AutoSize = true;
            this.chbxFemale.Location = new System.Drawing.Point(201, 69);
            this.chbxFemale.Name = "chbxFemale";
            this.chbxFemale.Size = new System.Drawing.Size(73, 17);
            this.chbxFemale.TabIndex = 11;
            this.chbxFemale.Text = "Женский";
            this.chbxFemale.UseVisualStyleBackColor = true;
            this.chbxFemale.CheckedChanged += new System.EventHandler(this.chbxFemale_CheckedChanged);
            // 
            // chlbxAnalysis
            // 
            this.chlbxAnalysis.CheckOnClick = true;
            this.chlbxAnalysis.FormattingEnabled = true;
            this.chlbxAnalysis.Location = new System.Drawing.Point(352, 53);
            this.chlbxAnalysis.Name = "chlbxAnalysis";
            this.chlbxAnalysis.Size = new System.Drawing.Size(200, 139);
            this.chlbxAnalysis.TabIndex = 12;
            this.chlbxAnalysis.SelectedIndexChanged += new System.EventHandler(this.chlbxAnalysis_SelectedIndexChanged);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.Location = new System.Drawing.Point(352, 10);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(200, 37);
            this.textBox1.TabIndex = 14;
            this.textBox1.Text = "Выберите виды анализа, включаемые в отчет";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // excelButton
            // 
            this.excelButton.Location = new System.Drawing.Point(128, 216);
            this.excelButton.Name = "excelButton";
            this.excelButton.Size = new System.Drawing.Size(82, 36);
            this.excelButton.TabIndex = 15;
            this.excelButton.Text = "Excel";
            this.excelButton.UseVisualStyleBackColor = true;
            this.excelButton.Visible = false;
            this.excelButton.Click += new System.EventHandler(this.excelButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 270);
            this.Controls.Add(this.excelButton);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.chlbxAnalysis);
            this.Controls.Add(this.chbxFemale);
            this.Controls.Add(this.chbxMale);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtPicker);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbxLabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbxName);
            this.Name = "MainForm";
            this.Text = "Отчет";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbxName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbxLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtPicker;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chbxMale;
        private System.Windows.Forms.CheckBox chbxFemale;
        private System.Windows.Forms.CheckedListBox chlbxAnalysis;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button excelButton;
    }
}

