namespace Laporan_Automation
{
    partial class Home
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
            this.RL53 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // RL53
            // 
            this.RL53.Location = new System.Drawing.Point(12, 12);
            this.RL53.Name = "RL53";
            this.RL53.Size = new System.Drawing.Size(227, 38);
            this.RL53.TabIndex = 2;
            this.RL53.Text = "RL 5.3";
            this.RL53.UseVisualStyleBackColor = true;
            this.RL53.Click += new System.EventHandler(this.RL53_Click);
            // 
            // Home
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(251, 62);
            this.Controls.Add(this.RL53);
            this.Name = "Home";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Home";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button RL53;
    }
}

