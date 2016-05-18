namespace settingService
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.butSetService = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lbServiceStatus = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 181F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 146F));
            this.tableLayoutPanel1.Controls.Add(this.button4, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.button3, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.button2, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.butSetService, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(47, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(655, 100);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button4.Location = new System.Drawing.Point(512, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(140, 94);
            this.button4.TabIndex = 3;
            this.button4.Text = "啟動服務";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button3.Location = new System.Drawing.Point(331, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(146, 94);
            this.button3.TabIndex = 2;
            this.button3.Text = "暫停服務";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.Location = new System.Drawing.Point(167, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(144, 94);
            this.button2.TabIndex = 1;
            this.button2.Text = "移除";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // butSetService
            // 
            this.butSetService.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.butSetService.Location = new System.Drawing.Point(3, 3);
            this.butSetService.Name = "butSetService";
            this.butSetService.Size = new System.Drawing.Size(137, 94);
            this.butSetService.TabIndex = 0;
            this.butSetService.Text = "安裝服務";
            this.butSetService.UseVisualStyleBackColor = true;
            this.butSetService.Click += new System.EventHandler(this.butSetService_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(44, 130);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 31);
            this.label1.TabIndex = 1;
            this.label1.Text = "服務:";
            // 
            // lbServiceStatus
            // 
            this.lbServiceStatus.AutoSize = true;
            this.lbServiceStatus.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbServiceStatus.Location = new System.Drawing.Point(136, 130);
            this.lbServiceStatus.Name = "lbServiceStatus";
            this.lbServiceStatus.Size = new System.Drawing.Size(62, 31);
            this.lbServiceStatus.TabIndex = 2;
            this.lbServiceStatus.Text = "狀態";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(727, 176);
            this.Controls.Add(this.lbServiceStatus);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button butSetService;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbServiceStatus;
    }
}