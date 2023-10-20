using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelImportExport
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
            this.ImportBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.ResultsList = new System.Windows.Forms.ListBox();
            this.ResultsList2 = new System.Windows.Forms.ListBox();
            this.coincidences = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.coincidencesBox = new System.Windows.Forms.ListBox();
            this.coincidencesLb = new System.Windows.Forms.Label();
            this.missingBox = new System.Windows.Forms.ListBox();
            this.missingLb = new System.Windows.Forms.Label();
            this.missingCountLb = new System.Windows.Forms.Label();
            this.remarksCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ImportBtn
            // 
            this.ImportBtn.Cursor = System.Windows.Forms.Cursors.Cross;
            this.ImportBtn.Location = new System.Drawing.Point(41, 32);
            this.ImportBtn.Name = "ImportBtn";
            this.ImportBtn.Size = new System.Drawing.Size(91, 34);
            this.ImportBtn.TabIndex = 0;
            this.ImportBtn.Text = "Import From Excel";
            this.ImportBtn.UseVisualStyleBackColor = true;
            this.ImportBtn.Click += new System.EventHandler(this.ImportBtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ResultsList
            // 
            this.ResultsList.FormattingEnabled = true;
            this.ResultsList.Location = new System.Drawing.Point(210, 32);
            this.ResultsList.Name = "ResultsList";
            this.ResultsList.Size = new System.Drawing.Size(102, 186);
            this.ResultsList.TabIndex = 1;
            // 
            // ResultsList2
            // 
            this.ResultsList2.FormattingEnabled = true;
            this.ResultsList2.Location = new System.Drawing.Point(12, 228);
            this.ResultsList2.Name = "ResultsList2";
            this.ResultsList2.Size = new System.Drawing.Size(159, 173);
            this.ResultsList2.TabIndex = 2;
            // 
            // coincidences
            // 
            this.coincidences.AutoSize = true;
            this.coincidences.Location = new System.Drawing.Point(478, 11);
            this.coincidences.Name = "coincidences";
            this.coincidences.Size = new System.Drawing.Size(13, 13);
            this.coincidences.TabIndex = 3;
            this.coincidences.Text = "0";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(217, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Remarks Id";
            // 
            // coincidencesBox
            // 
            this.coincidencesBox.FormattingEnabled = true;
            this.coincidencesBox.Location = new System.Drawing.Point(384, 32);
            this.coincidencesBox.Name = "coincidencesBox";
            this.coincidencesBox.Size = new System.Drawing.Size(120, 186);
            this.coincidencesBox.TabIndex = 5;
            // 
            // coincidencesLb
            // 
            this.coincidencesLb.AutoSize = true;
            this.coincidencesLb.Location = new System.Drawing.Point(384, 11);
            this.coincidencesLb.Name = "coincidencesLb";
            this.coincidencesLb.Size = new System.Drawing.Size(71, 13);
            this.coincidencesLb.TabIndex = 6;
            this.coincidencesLb.Text = "Coincidences";
            // 
            // missingBox
            // 
            this.missingBox.FormattingEnabled = true;
            this.missingBox.Location = new System.Drawing.Point(564, 32);
            this.missingBox.Name = "missingBox";
            this.missingBox.Size = new System.Drawing.Size(120, 186);
            this.missingBox.TabIndex = 7;
            // 
            // missingLb
            // 
            this.missingLb.AutoSize = true;
            this.missingLb.Location = new System.Drawing.Point(561, 11);
            this.missingLb.Name = "missingLb";
            this.missingLb.Size = new System.Drawing.Size(42, 13);
            this.missingLb.TabIndex = 6;
            this.missingLb.Text = "Missing";
            // 
            // missingCountLb
            // 
            this.missingCountLb.AutoSize = true;
            this.missingCountLb.Location = new System.Drawing.Point(618, 11);
            this.missingCountLb.Name = "missingCountLb";
            this.missingCountLb.Size = new System.Drawing.Size(13, 13);
            this.missingCountLb.TabIndex = 8;
            this.missingCountLb.Text = "0";
            // 
            // remarksCount
            // 
            this.remarksCount.AutoSize = true;
            this.remarksCount.Location = new System.Drawing.Point(284, 11);
            this.remarksCount.Name = "remarksCount";
            this.remarksCount.Size = new System.Drawing.Size(13, 13);
            this.remarksCount.TabIndex = 8;
            this.remarksCount.Text = "0";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.remarksCount);
            this.Controls.Add(this.missingCountLb);
            this.Controls.Add(this.missingBox);
            this.Controls.Add(this.missingLb);
            this.Controls.Add(this.coincidencesLb);
            this.Controls.Add(this.coincidencesBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.coincidences);
            this.Controls.Add(this.ResultsList2);
            this.Controls.Add(this.ResultsList);
            this.Controls.Add(this.ImportBtn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ImportBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ListBox ResultsList;
        private ListBox ResultsList2;
        private Label coincidences;
        private Label label1;
        private ListBox coincidencesBox;
        private Label coincidencesLb;
        private ListBox missingBox;
        private Label missingLb;
        private Label missingCountLb;
        private Label remarksCount;
    }
}

