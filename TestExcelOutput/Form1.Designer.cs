namespace TestExcelOutput
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnTestExcelOutput = new System.Windows.Forms.Button();
            this.saveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnTestExcelOutput
            // 
            this.btnTestExcelOutput.Location = new System.Drawing.Point(411, 470);
            this.btnTestExcelOutput.Name = "btnTestExcelOutput";
            this.btnTestExcelOutput.Size = new System.Drawing.Size(175, 53);
            this.btnTestExcelOutput.TabIndex = 0;
            this.btnTestExcelOutput.Text = "テスト帳票";
            this.btnTestExcelOutput.UseVisualStyleBackColor = true;
            this.btnTestExcelOutput.Click += new System.EventHandler(this.btnTestExcelOutput_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(50, 56);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(245, 64);
            this.button1.TabIndex = 1;
            this.button1.Text = "①一般衛生管理_清掃記録";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(50, 139);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(245, 63);
            this.button2.TabIndex = 2;
            this.button2.Text = "②一般衛生管理_実施記録";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1030, 587);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnTestExcelOutput);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnTestExcelOutput;
        private System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

