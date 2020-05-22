namespace SynEstTool
{
    partial class SheetSelector
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
            this.Select_No = new System.Windows.Forms.ListBox();
            this.Select_Yes = new System.Windows.Forms.ListBox();
            this.BtnItemMoveRight = new System.Windows.Forms.Button();
            this.BtnItemMoveLeft = new System.Windows.Forms.Button();
            this.BtnConsolidate = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Select_No
            // 
            this.Select_No.FormattingEnabled = true;
            this.Select_No.Location = new System.Drawing.Point(11, 11);
            this.Select_No.Margin = new System.Windows.Forms.Padding(2);
            this.Select_No.Name = "Select_No";
            this.Select_No.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.Select_No.Size = new System.Drawing.Size(150, 290);
            this.Select_No.TabIndex = 0;
            // 
            // Select_Yes
            // 
            this.Select_Yes.FormattingEnabled = true;
            this.Select_Yes.Location = new System.Drawing.Point(211, 11);
            this.Select_Yes.Margin = new System.Windows.Forms.Padding(2);
            this.Select_Yes.Name = "Select_Yes";
            this.Select_Yes.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.Select_Yes.Size = new System.Drawing.Size(150, 290);
            this.Select_Yes.TabIndex = 1;
            // 
            // BtnItemMoveRight
            // 
            this.BtnItemMoveRight.Location = new System.Drawing.Point(166, 102);
            this.BtnItemMoveRight.Name = "BtnItemMoveRight";
            this.BtnItemMoveRight.Size = new System.Drawing.Size(40, 40);
            this.BtnItemMoveRight.TabIndex = 2;
            this.BtnItemMoveRight.Text = ">>";
            this.BtnItemMoveRight.UseVisualStyleBackColor = true;
            this.BtnItemMoveRight.Click += new System.EventHandler(this.BtnItemMoveRight_Click);
            // 
            // BtnItemMoveLeft
            // 
            this.BtnItemMoveLeft.Location = new System.Drawing.Point(166, 148);
            this.BtnItemMoveLeft.Name = "BtnItemMoveLeft";
            this.BtnItemMoveLeft.Size = new System.Drawing.Size(40, 40);
            this.BtnItemMoveLeft.TabIndex = 3;
            this.BtnItemMoveLeft.Text = "<<";
            this.BtnItemMoveLeft.UseVisualStyleBackColor = true;
            this.BtnItemMoveLeft.Click += new System.EventHandler(this.BtnItemMoveLeft_Click);
            // 
            // BtnConsolidate
            // 
            this.BtnConsolidate.Location = new System.Drawing.Point(13, 306);
            this.BtnConsolidate.Name = "BtnConsolidate";
            this.BtnConsolidate.Size = new System.Drawing.Size(268, 39);
            this.BtnConsolidate.TabIndex = 4;
            this.BtnConsolidate.Text = "Consolidate Selected Sheets";
            this.BtnConsolidate.UseVisualStyleBackColor = true;
            this.BtnConsolidate.Click += new System.EventHandler(this.BtnConsolidate_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(287, 306);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(73, 39);
            this.BtnCancel.TabIndex = 5;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // SheetSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(372, 358);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnConsolidate);
            this.Controls.Add(this.BtnItemMoveLeft);
            this.Controls.Add(this.BtnItemMoveRight);
            this.Controls.Add(this.Select_Yes);
            this.Controls.Add(this.Select_No);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "SheetSelector";
            this.Text = "SheetSelector";
            this.Load += new System.EventHandler(this.SheetSelector_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox Select_No;
        private System.Windows.Forms.ListBox Select_Yes;
        private System.Windows.Forms.Button BtnItemMoveRight;
        private System.Windows.Forms.Button BtnItemMoveLeft;
        private System.Windows.Forms.Button BtnConsolidate;
        private System.Windows.Forms.Button BtnCancel;
    }
}