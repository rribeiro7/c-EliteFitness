namespace EliteFitness
{
    partial class Popup
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lbContato = new System.Windows.Forms.Label();
            this.lbInstrutor = new System.Windows.Forms.Label();
            this.lbDays = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(400, 400);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // lbContato
            // 
            this.lbContato.AutoSize = true;
            this.lbContato.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbContato.Location = new System.Drawing.Point(13, 419);
            this.lbContato.Name = "lbContato";
            this.lbContato.Size = new System.Drawing.Size(0, 16);
            this.lbContato.TabIndex = 1;
            // 
            // lbInstrutor
            // 
            this.lbInstrutor.AutoSize = true;
            this.lbInstrutor.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbInstrutor.Location = new System.Drawing.Point(13, 432);
            this.lbInstrutor.Name = "lbInstrutor";
            this.lbInstrutor.Size = new System.Drawing.Size(0, 16);
            this.lbInstrutor.TabIndex = 2;
            // 
            // lbDays
            // 
            this.lbDays.AutoSize = true;
            this.lbDays.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbDays.Location = new System.Drawing.Point(13, 459);
            this.lbDays.Name = "lbDays";
            this.lbDays.Size = new System.Drawing.Size(0, 16);
            this.lbDays.TabIndex = 3;
            // 
            // Popup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(425, 509);
            this.Controls.Add(this.lbDays);
            this.Controls.Add(this.lbInstrutor);
            this.Controls.Add(this.lbContato);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Popup";
            this.Text = "Imagem";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lbContato;
        private System.Windows.Forms.Label lbInstrutor;
        private System.Windows.Forms.Label lbDays;
    }
}