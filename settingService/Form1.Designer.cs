namespace settingService
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.butSave = new System.Windows.Forms.Button();
            this.txtCooperPath = new System.Windows.Forms.TextBox();
            this.txtXmlFromPath = new System.Windows.Forms.TextBox();
            this.txtXmlToPath = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtStartTime = new System.Windows.Forms.TextBox();
            this.txtEndTime = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.txtInterVal = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(103, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "cooper的路徑";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(97, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(154, 26);
            this.label2.TabIndex = 1;
            this.label2.Text = "xml的來源路徑";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(76, 142);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(175, 26);
            this.label3.TabIndex = 2;
            this.label3.Text = "xml的搬移的路徑";
            // 
            // butSave
            // 
            this.butSave.BackgroundImage = global::settingService.Properties.Resources.button;
            this.butSave.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.butSave.Cursor = System.Windows.Forms.Cursors.Hand;
            this.butSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butSave.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.butSave.Location = new System.Drawing.Point(279, 367);
            this.butSave.Name = "butSave";
            this.butSave.Size = new System.Drawing.Size(201, 65);
            this.butSave.TabIndex = 4;
            this.butSave.Text = "儲存設定";
            this.butSave.UseVisualStyleBackColor = true;
            this.butSave.Click += new System.EventHandler(this.butSave_Click);
            // 
            // txtCooperPath
            // 
            this.txtCooperPath.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtCooperPath.Location = new System.Drawing.Point(279, 44);
            this.txtCooperPath.Name = "txtCooperPath";
            this.txtCooperPath.Size = new System.Drawing.Size(272, 27);
            this.txtCooperPath.TabIndex = 5;
            // 
            // txtXmlFromPath
            // 
            this.txtXmlFromPath.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtXmlFromPath.Location = new System.Drawing.Point(279, 97);
            this.txtXmlFromPath.Name = "txtXmlFromPath";
            this.txtXmlFromPath.Size = new System.Drawing.Size(272, 27);
            this.txtXmlFromPath.TabIndex = 6;
            // 
            // txtXmlToPath
            // 
            this.txtXmlToPath.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtXmlToPath.Location = new System.Drawing.Point(279, 148);
            this.txtXmlToPath.Name = "txtXmlToPath";
            this.txtXmlToPath.Size = new System.Drawing.Size(272, 27);
            this.txtXmlToPath.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label4.Font = new System.Drawing.Font("新細明體", 14F);
            this.label4.Location = new System.Drawing.Point(19, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 19);
            this.label4.TabIndex = 8;
            this.label4.Text = "開始時間";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label5.Font = new System.Drawing.Font("新細明體", 14F);
            this.label5.Location = new System.Drawing.Point(19, 72);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 19);
            this.label5.TabIndex = 9;
            this.label5.Text = "結束時間";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(88, 203);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(163, 26);
            this.label6.TabIndex = 10;
            this.label6.Text = "匯入XML的時間";
            // 
            // txtStartTime
            // 
            this.txtStartTime.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtStartTime.Location = new System.Drawing.Point(130, 21);
            this.txtStartTime.MaxLength = 4;
            this.txtStartTime.Name = "txtStartTime";
            this.txtStartTime.Size = new System.Drawing.Size(165, 27);
            this.txtStartTime.TabIndex = 11;
            // 
            // txtEndTime
            // 
            this.txtEndTime.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtEndTime.Location = new System.Drawing.Point(130, 64);
            this.txtEndTime.MaxLength = 4;
            this.txtEndTime.Name = "txtEndTime";
            this.txtEndTime.Size = new System.Drawing.Size(165, 27);
            this.txtEndTime.TabIndex = 12;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label7.Font = new System.Drawing.Font("新細明體", 14F);
            this.label7.ForeColor = System.Drawing.Color.Maroon;
            this.label7.Location = new System.Drawing.Point(305, 29);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(110, 19);
            this.label7.TabIndex = 13;
            this.label7.Text = "( 舉例:1000 )";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label8.Font = new System.Drawing.Font("新細明體", 14F);
            this.label8.ForeColor = System.Drawing.Color.Maroon;
            this.label8.Location = new System.Drawing.Point(305, 72);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(110, 19);
            this.label8.TabIndex = 14;
            this.label8.Text = "( 舉例:2400 )";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("新細明體", 12F);
            this.button1.Location = new System.Drawing.Point(574, 44);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(47, 35);
            this.button1.TabIndex = 15;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("新細明體", 12F);
            this.button2.Location = new System.Drawing.Point(574, 95);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(47, 35);
            this.button2.TabIndex = 16;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("新細明體", 12F);
            this.button3.Location = new System.Drawing.Point(574, 142);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(47, 35);
            this.button3.TabIndex = 17;
            this.button3.Text = "...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label9.Font = new System.Drawing.Font("新細明體", 14F);
            this.label9.Location = new System.Drawing.Point(19, 107);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(85, 19);
            this.label9.TabIndex = 18;
            this.label9.Text = "間隔時間";
            // 
            // txtInterVal
            // 
            this.txtInterVal.Font = new System.Drawing.Font("新細明體", 12F);
            this.txtInterVal.Location = new System.Drawing.Point(130, 106);
            this.txtInterVal.MaxLength = 4;
            this.txtInterVal.Name = "txtInterVal";
            this.txtInterVal.Size = new System.Drawing.Size(71, 27);
            this.txtInterVal.TabIndex = 19;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label10.Font = new System.Drawing.Font("新細明體", 14F);
            this.label10.ForeColor = System.Drawing.Color.Maroon;
            this.label10.Location = new System.Drawing.Point(426, 274);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(47, 19);
            this.label10.TabIndex = 20;
            this.label10.Text = "分鐘";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtInterVal);
            this.groupBox1.Controls.Add(this.txtStartTime);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtEndTime);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Location = new System.Drawing.Point(279, 203);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(433, 146);
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.label11.Font = new System.Drawing.Font("新細明體", 14F);
            this.label11.ForeColor = System.Drawing.Color.Maroon;
            this.label11.Location = new System.Drawing.Point(221, 110);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(47, 19);
            this.label11.TabIndex = 20;
            this.label11.Text = "分鐘";
            // 
            // button4
            // 
            this.button4.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button4.BackgroundImage")));
            this.button4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button4.Location = new System.Drawing.Point(505, 367);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(207, 65);
            this.button4.TabIndex = 22;
            this.button4.Text = "設定服務";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::settingService.Properties.Resources.background;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(761, 479);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtXmlToPath);
            this.Controls.Add(this.txtXmlFromPath);
            this.Controls.Add(this.txtCooperPath);
            this.Controls.Add(this.butSave);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CooperService";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button butSave;
        private System.Windows.Forms.TextBox txtCooperPath;
        private System.Windows.Forms.TextBox txtXmlFromPath;
        private System.Windows.Forms.TextBox txtXmlToPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtStartTime;
        private System.Windows.Forms.TextBox txtEndTime;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtInterVal;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button button4;
    }
}

