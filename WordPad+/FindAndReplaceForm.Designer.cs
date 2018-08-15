namespace WordPad_
{
    partial class FindAndReplaceForm
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
            this.findWhatLabel = new System.Windows.Forms.Label();
            this.findTB = new System.Windows.Forms.TextBox();
            this.findNextButton = new System.Windows.Forms.Button();
            this.caseSensitiveCB = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // findWhatLabel
            // 
            this.findWhatLabel.AutoSize = true;
            this.findWhatLabel.Location = new System.Drawing.Point(12, 15);
            this.findWhatLabel.Name = "findWhatLabel";
            this.findWhatLabel.Size = new System.Drawing.Size(35, 13);
            this.findWhatLabel.TabIndex = 0;
            this.findWhatLabel.Text = "label1";
            // 
            // findTB
            // 
            this.findTB.Location = new System.Drawing.Point(74, 12);
            this.findTB.Name = "findTB";
            this.findTB.Size = new System.Drawing.Size(187, 20);
            this.findTB.TabIndex = 1;
            // 
            // findNextButton
            // 
            this.findNextButton.Location = new System.Drawing.Point(267, 10);
            this.findNextButton.Name = "findNextButton";
            this.findNextButton.Size = new System.Drawing.Size(75, 23);
            this.findNextButton.TabIndex = 2;
            this.findNextButton.Text = "Find next";
            this.findNextButton.UseVisualStyleBackColor = true;
            this.findNextButton.Click += new System.EventHandler(this.findNextButton_Click);
            // 
            // caseSensitiveCB
            // 
            this.caseSensitiveCB.AutoSize = true;
            this.caseSensitiveCB.Location = new System.Drawing.Point(12, 68);
            this.caseSensitiveCB.Name = "caseSensitiveCB";
            this.caseSensitiveCB.Size = new System.Drawing.Size(94, 17);
            this.caseSensitiveCB.TabIndex = 3;
            this.caseSensitiveCB.Text = "Case sensitive";
            this.caseSensitiveCB.UseVisualStyleBackColor = true;
            // 
            // FindAndReplaceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(354, 97);
            this.Controls.Add(this.caseSensitiveCB);
            this.Controls.Add(this.findNextButton);
            this.Controls.Add(this.findTB);
            this.Controls.Add(this.findWhatLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FindAndReplaceForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FindAndReplace";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label findWhatLabel;
        private System.Windows.Forms.TextBox findTB;
        private System.Windows.Forms.Button findNextButton;
        private System.Windows.Forms.CheckBox caseSensitiveCB;
    }
}