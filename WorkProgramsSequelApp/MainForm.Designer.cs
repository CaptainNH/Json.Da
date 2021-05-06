
namespace WorkProgramsSequelApp
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.buttonDB = new System.Windows.Forms.Button();
            this.buttonGenerateDocs = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Font = new System.Drawing.Font("Britannic Bold", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(561, 120);
            this.label1.TabIndex = 0;
            this.label1.Text = "MathFuck";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // buttonDB
            // 
            this.buttonDB.BackColor = System.Drawing.Color.SteelBlue;
            this.buttonDB.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonDB.Font = new System.Drawing.Font("Britannic Bold", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.buttonDB.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonDB.Location = new System.Drawing.Point(92, 234);
            this.buttonDB.Name = "buttonDB";
            this.buttonDB.Size = new System.Drawing.Size(362, 68);
            this.buttonDB.TabIndex = 1;
            this.buttonDB.Text = "База данных";
            this.buttonDB.UseVisualStyleBackColor = false;
            // 
            // buttonGenerateDocs
            // 
            this.buttonGenerateDocs.BackColor = System.Drawing.Color.SteelBlue;
            this.buttonGenerateDocs.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonGenerateDocs.Font = new System.Drawing.Font("Britannic Bold", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.buttonGenerateDocs.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.buttonGenerateDocs.Location = new System.Drawing.Point(92, 366);
            this.buttonGenerateDocs.Name = "buttonGenerateDocs";
            this.buttonGenerateDocs.Size = new System.Drawing.Size(362, 68);
            this.buttonGenerateDocs.TabIndex = 2;
            this.buttonGenerateDocs.Text = "Генерация документов";
            this.buttonGenerateDocs.UseVisualStyleBackColor = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(561, 571);
            this.Controls.Add(this.buttonGenerateDocs);
            this.Controls.Add(this.buttonDB);
            this.Controls.Add(this.label1);
            this.Name = "MainForm";
            this.Text = "Рабочие программы&Co.";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonDB;
        private System.Windows.Forms.Button buttonGenerateDocs;
    }
}

