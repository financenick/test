namespace WF1
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.browseRkkFileButton = new System.Windows.Forms.Button();
            this.browseRequestsFileButton = new System.Windows.Forms.Button();
            this.rkkFileTextBox = new System.Windows.Forms.TextBox();
            this.requestsFileTextBox = new System.Windows.Forms.TextBox();
            this.calculateButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // browseRkkFileButton
            // 
            this.browseRkkFileButton.Location = new System.Drawing.Point(43, 50);
            this.browseRkkFileButton.Name = "browseRkkFileButton";
            this.browseRkkFileButton.Size = new System.Drawing.Size(173, 61);
            this.browseRkkFileButton.TabIndex = 2;
            this.browseRkkFileButton.Text = "Данные по РКК";
            this.browseRkkFileButton.UseVisualStyleBackColor = true;
            this.browseRkkFileButton.Click += new System.EventHandler(this.browseRkkFileButton_Click);
            // 
            // browseRequestsFileButton
            // 
            this.browseRequestsFileButton.Location = new System.Drawing.Point(260, 50);
            this.browseRequestsFileButton.Name = "browseRequestsFileButton";
            this.browseRequestsFileButton.Size = new System.Drawing.Size(173, 61);
            this.browseRequestsFileButton.TabIndex = 3;
            this.browseRequestsFileButton.Text = "Данные по обращениям";
            this.browseRequestsFileButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.browseRequestsFileButton.UseVisualStyleBackColor = true;
            this.browseRequestsFileButton.Click += new System.EventHandler(this.browseRequestsFileButton_Click);
            // 
            // rkkFileTextBox
            // 
            this.rkkFileTextBox.Location = new System.Drawing.Point(43, 117);
            this.rkkFileTextBox.Name = "rkkFileTextBox";
            this.rkkFileTextBox.Size = new System.Drawing.Size(173, 22);
            this.rkkFileTextBox.TabIndex = 4;
            this.rkkFileTextBox.TextChanged += new System.EventHandler(this.rkkFileTextBox_TextChanged);
            // 
            // requestsFileTextBox
            // 
            this.requestsFileTextBox.Location = new System.Drawing.Point(260, 117);
            this.requestsFileTextBox.Name = "requestsFileTextBox";
            this.requestsFileTextBox.Size = new System.Drawing.Size(173, 22);
            this.requestsFileTextBox.TabIndex = 5;
            this.requestsFileTextBox.TextChanged += new System.EventHandler(this.requestsFileTextBox_TextChanged);
            // 
            // calculateButton
            // 
            this.calculateButton.Location = new System.Drawing.Point(43, 196);
            this.calculateButton.Name = "calculateButton";
            this.calculateButton.Size = new System.Drawing.Size(173, 61);
            this.calculateButton.TabIndex = 6;
            this.calculateButton.Text = "Рассчитать";
            this.calculateButton.UseVisualStyleBackColor = true;
            this.calculateButton.Click += new System.EventHandler(this.calculateButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1011, 598);
            this.Controls.Add(this.calculateButton);
            this.Controls.Add(this.requestsFileTextBox);
            this.Controls.Add(this.rkkFileTextBox);
            this.Controls.Add(this.browseRequestsFileButton);
            this.Controls.Add(this.browseRkkFileButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button browseRkkFileButton;
        private System.Windows.Forms.Button browseRequestsFileButton;
        private System.Windows.Forms.TextBox rkkFileTextBox;
        private System.Windows.Forms.TextBox requestsFileTextBox;
        private System.Windows.Forms.Button calculateButton;
    }
}

