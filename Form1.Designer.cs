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
            this.ResultsList.Size = new System.Drawing.Size(264, 173);
            this.ResultsList.TabIndex = 1;
            // 
            // ResultsList2
            // 
            this.ResultsList2.FormattingEnabled = true;
            this.ResultsList2.Location = new System.Drawing.Point(12, 228);
            this.ResultsList2.Name = "ResultsList2";
            this.ResultsList2.Size = new System.Drawing.Size(764, 173);
            this.ResultsList2.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.ResultsList2);
            this.Controls.Add(this.ResultsList);
            this.Controls.Add(this.ImportBtn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ImportBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ListBox ResultsList;
        private ListBox ResultsList2;
    }
}

