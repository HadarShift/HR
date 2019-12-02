namespace HR
{
    partial class UpdateVersion
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
            this.lbl_NumEmpUp = new System.Windows.Forms.Label();
            this.lbl_password = new System.Windows.Forms.Label();
            this.txt_NumEmpUp = new System.Windows.Forms.TextBox();
            this.txt_Password = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_NumEmpUp
            // 
            this.lbl_NumEmpUp.AutoSize = true;
            this.lbl_NumEmpUp.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_NumEmpUp.Location = new System.Drawing.Point(279, 26);
            this.lbl_NumEmpUp.Name = "lbl_NumEmpUp";
            this.lbl_NumEmpUp.Size = new System.Drawing.Size(102, 24);
            this.lbl_NumEmpUp.TabIndex = 0;
            this.lbl_NumEmpUp.Text = "מספר עובד";
            // 
            // lbl_password
            // 
            this.lbl_password.AutoSize = true;
            this.lbl_password.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_password.Location = new System.Drawing.Point(316, 69);
            this.lbl_password.Name = "lbl_password";
            this.lbl_password.Size = new System.Drawing.Size(65, 24);
            this.lbl_password.TabIndex = 1;
            this.lbl_password.Text = "סיסמא";
            // 
            // txt_NumEmpUp
            // 
            this.txt_NumEmpUp.Location = new System.Drawing.Point(78, 30);
            this.txt_NumEmpUp.Name = "txt_NumEmpUp";
            this.txt_NumEmpUp.Size = new System.Drawing.Size(182, 20);
            this.txt_NumEmpUp.TabIndex = 2;
            // 
            // txt_Password
            // 
            this.txt_Password.Location = new System.Drawing.Point(78, 69);
            this.txt_Password.Name = "txt_Password";
            this.txt_Password.Size = new System.Drawing.Size(182, 20);
            this.txt_Password.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.button1.Location = new System.Drawing.Point(201, 120);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 34);
            this.button1.TabIndex = 4;
            this.button1.Text = "אשר";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // UpdateVersion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(459, 166);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txt_Password);
            this.Controls.Add(this.txt_NumEmpUp);
            this.Controls.Add(this.lbl_password);
            this.Controls.Add(this.lbl_NumEmpUp);
            this.Name = "UpdateVersion";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UpdateVersion";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_NumEmpUp;
        private System.Windows.Forms.Label lbl_password;
        private System.Windows.Forms.TextBox txt_NumEmpUp;
        private System.Windows.Forms.TextBox txt_Password;
        private System.Windows.Forms.Button button1;
    }
}