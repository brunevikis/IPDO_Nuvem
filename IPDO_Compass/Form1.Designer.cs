namespace IPDO_Compass
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
            this.txt_Caminho = new System.Windows.Forms.TextBox();
            this.bt_seleciona = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txt_Caminho
            // 
            this.txt_Caminho.Location = new System.Drawing.Point(12, 46);
            this.txt_Caminho.Name = "txt_Caminho";
            this.txt_Caminho.Size = new System.Drawing.Size(427, 20);
            this.txt_Caminho.TabIndex = 0;
            // 
            // bt_seleciona
            // 
            this.bt_seleciona.Location = new System.Drawing.Point(445, 44);
            this.bt_seleciona.Name = "bt_seleciona";
            this.bt_seleciona.Size = new System.Drawing.Size(75, 23);
            this.bt_seleciona.TabIndex = 1;
            this.bt_seleciona.Text = "Selecionar";
            this.bt_seleciona.UseVisualStyleBackColor = true;
            this.bt_seleciona.Click += new System.EventHandler(this.bt_seleciona_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Caminho Arquivo IPDO:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bt_seleciona);
            this.Controls.Add(this.txt_Caminho);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_Caminho;
        private System.Windows.Forms.Button bt_seleciona;
        private System.Windows.Forms.Label label1;
    }
}

