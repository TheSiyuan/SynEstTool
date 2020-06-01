namespace SynEstTool
{
    partial class ReviewCriteria
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.NoReview = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FeeReview = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FullReview = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.NoReview,
            this.FeeReview,
            this.FullReview});
            this.dataGridView1.Location = new System.Drawing.Point(63, 47);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(619, 450);
            this.dataGridView1.TabIndex = 0;
            // 
            // NoReview
            // 
            this.NoReview.HeaderText = "No Review";
            this.NoReview.MinimumWidth = 8;
            this.NoReview.Name = "NoReview";
            this.NoReview.Width = 150;
            // 
            // FeeReview
            // 
            this.FeeReview.HeaderText = "Fee Review";
            this.FeeReview.MinimumWidth = 8;
            this.FeeReview.Name = "FeeReview";
            this.FeeReview.Width = 150;
            // 
            // FullReview
            // 
            this.FullReview.HeaderText = "FullReview";
            this.FullReview.MinimumWidth = 8;
            this.FullReview.Name = "FullReview";
            this.FullReview.Width = 150;
            // 
            // ReviewCriteria
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1199, 710);
            this.Controls.Add(this.dataGridView1);
            this.Name = "ReviewCriteria";
            this.Text = "ReviewCriteria";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn NoReview;
        private System.Windows.Forms.DataGridViewTextBoxColumn FeeReview;
        private System.Windows.Forms.DataGridViewTextBoxColumn FullReview;
    }
}