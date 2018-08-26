namespace VisioStatusUML
{
    partial class frmXMLVisio
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
            this.btnVisio = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnVisio
            // 
            this.btnVisio.Location = new System.Drawing.Point(12, 12);
            this.btnVisio.Name = "btnVisio";
            this.btnVisio.Size = new System.Drawing.Size(103, 41);
            this.btnVisio.TabIndex = 2;
            this.btnVisio.Text = "Gerar";
            this.btnVisio.UseVisualStyleBackColor = true;
            this.btnVisio.Click += new System.EventHandler(this.btnVisio_Click);
            // 
            // frmXMLVisio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(127, 65);
            this.Controls.Add(this.btnVisio);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmXMLVisio";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UML Visio";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnVisio;
    }
}

