namespace AppsLaunchTime
{
    partial class Form1
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
            button1 = new Button();
            textBox1 = new TextBox();
            button2 = new Button();
            numericUpDownIterations = new NumericUpDown();
            labelIterations = new Label();
            ((System.ComponentModel.ISupportInitialize)numericUpDownIterations).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(1, 12);
            button1.Name = "button1";
            button1.Size = new Size(69, 34);
            button1.TabIndex = 0;
            button1.Text = "Start";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(1, 52);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(457, 180);
            textBox1.TabIndex = 1;
            textBox1.Multiline = true;
            textBox1.ReadOnly = true;
            textBox1.Font = new Font("Consolas", 9F);
            textBox1.ScrollBars = ScrollBars.Vertical;
            // 
            // button2
            // 
            button2.Location = new Point(87, 12);
            button2.Name = "button2";
            button2.Size = new Size(61, 34);
            button2.TabIndex = 2;
            button2.Text = "Camera";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // labelIterations
            // 
            labelIterations.AutoSize = true;
            labelIterations.Location = new Point(160, 20);
            labelIterations.Name = "labelIterations";
            labelIterations.Text = "Iterations:";
            // 
            // numericUpDownIterations
            // 
            numericUpDownIterations.Location = new Point(260, 17);
            numericUpDownIterations.Name = "numericUpDownIterations";
            numericUpDownIterations.Size = new Size(70, 31);
            numericUpDownIterations.Minimum = 1;
            numericUpDownIterations.Maximum = 1000;
            numericUpDownIterations.Value = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(457, 240);
            Controls.Add(button2);
            Controls.Add(textBox1);
            Controls.Add(button1);
            Controls.Add(labelIterations);
            Controls.Add(numericUpDownIterations);
            Name = "Form1";
            Text = "AppslaunchTime";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)numericUpDownIterations).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private TextBox textBox1;
        private Button button2;
        private NumericUpDown numericUpDownIterations;
        private Label labelIterations;
    }
}
