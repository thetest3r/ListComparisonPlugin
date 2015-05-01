namespace ListProcessingExcelPlugin
{
    partial class EasterEgg
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
            this.step1 = new System.Windows.Forms.Label();
            this.step2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // step1
            // 
            this.step1.AutoSize = true;
            this.step1.Location = new System.Drawing.Point(69, 9);
            this.step1.Name = "step1";
            this.step1.Size = new System.Drawing.Size(499, 13);
            this.step1.TabIndex = 2;
            this.step1.Text = "1. Notice the difference between the two lists below: (one has two fields combine" +
    "d, the other is \"normal\")";
            // 
            // step2
            // 
            this.step2.AutoSize = true;
            this.step2.Location = new System.Drawing.Point(72, 194);
            this.step2.Name = "step2";
            this.step2.Size = new System.Drawing.Size(433, 13);
            this.step2.TabIndex = 4;
            this.step2.Text = "2. There\'s a shortcut to this (without the need of using Excel to separate them b" +
    "y commas)!";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(72, 304);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(436, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Notice how the sheet on the left has an extra column, while the sheet on the righ" +
    "t does not.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(220, 319);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(166, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Notice the order [Last, First, Date]";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 363);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(543, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "This is the hidden easter egg! The plug-in should compare those two lists without" +
    " the need for separating commas!";
            // 
            // pictureBox4
            // 
            this.pictureBox4.BackgroundImage = global::ListProcessingExcelPlugin.Properties.Resources.ResultsList;
            this.pictureBox4.Location = new System.Drawing.Point(72, 388);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(453, 185);
            this.pictureBox4.TabIndex = 9;
            this.pictureBox4.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.BackgroundImage = global::ListProcessingExcelPlugin.Properties.Resources.PluginSettings2;
            this.pictureBox3.Location = new System.Drawing.Point(146, 219);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(313, 82);
            this.pictureBox3.TabIndex = 5;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackgroundImage = global::ListProcessingExcelPlugin.Properties.Resources.NormalSheet;
            this.pictureBox2.Location = new System.Drawing.Point(12, 34);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(314, 148);
            this.pictureBox2.TabIndex = 3;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::ListProcessingExcelPlugin.Properties.Resources.CommaSepSheet;
            this.pictureBox1.Location = new System.Drawing.Point(373, 34);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(276, 148);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // EasterEgg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(656, 590);
            this.Controls.Add(this.pictureBox4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox3);
            this.Controls.Add(this.step2);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.step1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "EasterEgg";
            this.Text = "EasterEgg";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label step1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label step2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBox4;
    }
}